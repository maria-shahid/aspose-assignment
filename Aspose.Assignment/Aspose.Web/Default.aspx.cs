using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aspose.Cells;
using System.Data;
using Aspose.Words;
using System.Web.Configuration;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;

namespace Aspose.Web
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            //Reset message label
            if (lblMessage.InnerText != "")
            {
                lblMessage.InnerText = "";

                if (lblMessage.Attributes["class"] != null)
                {
                    lblMessage.Attributes.Remove("class");
                }
            }
        }

        protected void btnFileUpload_Click(object sender, EventArgs e)
        {
            try
            {
                //Check if user attached a file
                if (employeeFile.HasFile)
                {
                    //Validate xlsx extension
                    if (!employeeFile.FileName.EndsWith(".xlsx"))
                    {
                        //Notify user to attach only xlsx file
                        lblMessage.InnerText = "ERROR: Only xlsx file type is allowed.";
                        lblMessage.Attributes.Add("class", "alert alert-danger");

                        //Show upload elements
                        paraAttach.Visible = true;

                        //hide the verify panel
                        paraVerify.Visible = false;
                        return;
                    }

                    // Instantiating a Workbook object
                    // Opening the Excel file through the file stream
                    Workbook workbook = new Workbook(employeeFile.PostedFile.InputStream);

                    // Accessing the first worksheet in the Excel file
                    Worksheet worksheet = workbook.Worksheets[0];
                    Boolean isValidFormat = true;                 

                    //Validate format
                    if ((worksheet.Cells.MaxColumn + 1) != 4 )
                    {
                        isValidFormat = false;                       
                    }
                    else
                    {
                        //Get Header cells from worksheet
                        Cell firstCell = worksheet.Cells[0, 0];
                        Cell secondCell = worksheet.Cells[0, 1];
                        Cell thirdCell = worksheet.Cells[0, 2];
                        Cell fourthCell = worksheet.Cells[0, 3];

                        //Check if valid header
                        if (firstCell.Value.ToString() != "FullName" || secondCell.Value.ToString() != "Email" || thirdCell.Value.ToString() != "Address" || fourthCell.Value.ToString() != "Salary")
                        {
                            isValidFormat = false;
                        }
                    }


                    if (isValidFormat)
                    {
                        // Exporting the contents of visible columns and rows starting from 1st cell to DataTable
                        DataTable dataTable = worksheet.Cells.ExportDataTable(0, 0, worksheet.Cells.MaxRow + 1, worksheet.Cells.MaxColumn + 1, true);

                        // Bind the grid
                        gvEmployees.DataSource = dataTable;
                        gvEmployees.DataBind();

                        //Save exported data table to ViewState for later use
                        ViewState["EmployeeData"] = dataTable;

                        //Show the verify grid and send button
                        paraVerify.Visible = true;

                        //Hide upload elements
                        paraAttach.Visible = false;
                    }
                    else {

                        //Notify user to attach a file with valid format
                        showMessage("ERROR: File format is not valid.", "alert alert-danger");
                                             
                        //Show upload elements
                        paraAttach.Visible = true;

                        //hide the verify panel
                        paraVerify.Visible = false;

                        return;
                    }
                  
                }
                else
                {
                    //Show upload elements
                    paraAttach.Visible = true;

                    //hide the verify grid and send button
                    paraVerify.Visible = false;

                    //Notify user to attach a file
                     showMessage("ERROR: Please attach a file to proceed.","alert alert-danger");
               
                }

            }
            catch (Exception ex)
            {
                showMessage("ERROR: An error occurred uploading file, please try again or contact support.", "alert alert-danger");     
            }
        }


        protected void showMessage(string message, string cssClass)
        {
           
            //Notify user 
            lblMessage.InnerText = message;

            if (lblMessage.Attributes["class"] != null)
            {
                lblMessage.Attributes["class"]= cssClass;
            }
            else
            {
                lblMessage.Attributes.Add("class", cssClass);
            }
           
        }

        protected void btnSend_Click(object sender, EventArgs e)
        {
            try
            {
                Boolean emailsSent = true;

                if (ViewState["EmployeeData"] != null)
                {
                    //Load datatable from ViewState
                    DataTable employeeData = (DataTable)ViewState["EmployeeData"];

                    //Init Email template path
                    string emailTemplate = Server.MapPath("~/" + WebConfigurationManager.AppSettings["TemplatesFolder"] + "/EmailTemplate.html");

                    //Init increment letter template
                    Document doc = new Document(Server.MapPath("~/" + WebConfigurationManager.AppSettings["TemplatesFolder"] + "/IncrementLetterTemplate.docx"));

                    // Loop through all records in the data source
                    foreach (DataRow row in employeeData.Rows)
                    {

                        if (row["Email"].ToString() != null && row["Email"].ToString() != "")
                        {
                            // Load email template
                            StreamReader reader = new StreamReader(emailTemplate);
                            String emailBody = reader.ReadToEnd();
                            emailBody = emailBody.Replace("{Name}", row["FullName"].ToString());

                            // Clone the template
                            Document letterTemplate = (Document)doc.Clone(true);

                            // Execute mail merge
                            letterTemplate.MailMerge.Execute(row);

                            // Save the document to memory stream
                            MemoryStream fileStream = new MemoryStream();
                            letterTemplate.Save(fileStream, Words.SaveFormat.Docx);
                            //Reset memory stream position
                            fileStream.Position = 0;

                            //Init email object and send email                          
                            SMTPEmail objEmail = new SMTPEmail();
                            objEmail.subject = "Increment Letter";
                            objEmail.toEmail = row["Email"].ToString();
                            objEmail.fileName = row["FullName"].ToString() + "-Increment-Letter.docx";
                            objEmail.fileStream = fileStream;
                            objEmail.body = emailBody;

                            if (!objEmail.SendHtmlEmail())
                            {
                                emailsSent = false;
                                break;
                            }
                        }


                    }

                    if (emailsSent)
                    {
                        //Hide the verify grid and send button
                        paraVerify.Visible = false;

                        //Show upload elements
                        paraAttach.Visible = true;

                        //Notify user after sending all emails
                        showMessage("SUCCESS: All emails sent successfully.", "alert alert-success");
                    }
                    else
                    {
                        //Notify user that email sending failed
                        showMessage("ERROR: An error occurred sending increment letters, please try again or contact support.", "alert alert-danger");
                    }

                }
                else
                {
                    //Hide the verify grid and send button
                    paraVerify.Visible = false;

                    //Show upload elements
                    paraAttach.Visible = true;

                    //Notify user to upload email template again
                    showMessage("ERROR: Failed to load employees data, please upload again.", "alert alert-danger");
                }
         
            }
            catch(Exception ex)
            {
                //Notify user that email sending failed
                showMessage("ERROR: An error occurred sending increment letters, please try again or contact support.","alert alert-danger");
            
            }
      
        }
    }
}