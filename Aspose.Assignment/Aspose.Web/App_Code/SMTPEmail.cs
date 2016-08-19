using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net.Mail;
using System.Web.Configuration;
using System.Net.Mime;
using System.IO;

namespace Aspose.Web
{
    public class SMTPEmail
    {
        public string subject;
        public string toEmail;
        public string errorMessage;
        public string body;
        public MemoryStream fileStream;
        public string fileName;
        public System.Net.Mime.ContentType contentType;
        public Attachment fileAttach;

        public bool SendHtmlEmail()
        {
            try
            {
              
                //Init mail message object
                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress(WebConfigurationManager.AppSettings["SMTPLogin"], "no-reply");
                mailMessage.Subject = subject;
                mailMessage.Body = body;
                mailMessage.IsBodyHtml = true;
                mailMessage.To.Add(new MailAddress(toEmail));
              
                //Add attachment if file stream and name exists
                if (fileStream != null && fileName != null && fileName != "")
                {
                    //Init file attachment
                    Attachment fileAttach = new Attachment(fileStream, fileName);

                    //Add attachment to mail message
                    mailMessage.Attachments.Add(fileAttach);
                }

                //Init smtp client
                SmtpClient smtp = new SmtpClient();
                smtp.Host = WebConfigurationManager.AppSettings["SMTPServer"];
                smtp.EnableSsl = true;
                System.Net.NetworkCredential NetworkCred = new System.Net.NetworkCredential();
                NetworkCred.UserName = WebConfigurationManager.AppSettings["SMTPLogin"];
                NetworkCred.Password = WebConfigurationManager.AppSettings["SMTPPassword"];
                smtp.UseDefaultCredentials = true;
                smtp.Credentials = NetworkCred;
                smtp.Port = int.Parse(WebConfigurationManager.AppSettings["SMTPPort"]);

                //Send Email
                smtp.Send(mailMessage);
               return true;

            }
            catch (SmtpException ex)
            {
                //catched smtp exception
                errorMessage = "SMTP ERROR: " + ex.Message.ToString();
                return  false;
            }
            catch (Exception ex)
            {
                errorMessage = "ERROR: " + ex.Message.ToString();
                return false;
            }

        }
    }
}