using System;
using System.IO;


namespace RASW
{
    public static class Logger
    {
        //public static void WriteToLog(string Data, string UserName);
        public static void WriteToLog(string Data)
        {
            try
            {
                string filename = Path.Combine(Directory.GetCurrentDirectory(),"WolseleyQuoteImport_Errors.txt"); //\\ak-inv-fs1\";

                    StreamWriter sr = new StreamWriter(filename, true);
                
                   // if(UserName.Length == 0) { UserName = "- - -"; }

                   if (Data != null) sr.WriteLine(DateTime.Now.ToString() + " -> " + Data); // + " -> User Name: " + UserName);
                    sr.Close();
                    sr.Dispose();

                   // MailMessage message = new MailMessage();
                   // message.To.Add("rwilson@akvagroup.com");
                   // message.Subject = "AKVA Connect Data Reporter..."; ;
                   // message.From = new System.Net.Mail.MailAddress("DataReporter@akvagroup.com");
                   // message.IsBodyHtml = false;
                   // message.Body = DateTime.Now.ToString() + " -> " + Data; // + " -> User Name: " + UserName;
                                  
                   // //Attachment data = new Attachment(filename);
                   //// message.Attachments.Add(data);
                    
                   // SmtpClient smtp = new SmtpClient("smtp.gmail.com");
                   // smtp.Send(message);
                   // message.Dispose();
                   // smtp.Dispose();

                //try
                //{
                //    MailMessage mail = new MailMessage();
                //    SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");

                //    mail.From = new MailAddress("datareporter@gmail.com");
                //    mail.To.Add("rwilson@akvagroup.net");
                //    mail.Subject = "Data Reporter Error...";
                //    mail.Body = DateTime.Now.ToString() + " -> " + Data + " -> User Name: "; // + UserName;
                //    mail.Attachments.Add(new Attachment(filename, "text/plain"));

                //    SmtpServer.Port = 587;
                //    SmtpServer.Credentials = new System.Net.NetworkCredential("routingsdocs@gmail.com", "x394yabr3");
                //    SmtpServer.EnableSsl = true;
                //    SmtpServer.Send(mail);
                //    SmtpServer.Dispose();
                //}
                //catch { }
            }
            catch  //(FileNotFoundException fileNotFound)
            {
            } 
        }
    }
}
