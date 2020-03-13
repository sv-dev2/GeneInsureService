using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace SmsService
{
   public class CustomerLogic
    {
        public async Task<string> SendSMS(string numberTO, string body)
        {
            Library.WriteErrorLog(numberTO);
            return "";
            using (var client = new HttpClient())
            {
                string username = System.Configuration.ConfigurationManager.AppSettings["smsGatewayUsername"].ToString();

                // Webservices token for above Webservice username
                string token = System.Configuration.ConfigurationManager.AppSettings["smsGatewayToken"].ToString();

                // BulkSMS Webservices URL
                string bulksms_ws = "http://portal.bulksmsweb.com/index.php?app=ws";

                // destination numbers, comma seperated or use #groupcode for sending to group
                // $destinations = '#devteam,263071077072,26370229338';
                // $destinations = '26300123123123,26300456456456';  for multiple recipients

                string destinations = numberTO;

                // SMS Message to send
                string message = body;

                // send via BulkSMS HTTP API

                string ws_str = bulksms_ws + "&u=" + username + "&h=" + token + "&op=pv";
                ws_str += "&to=" + Uri.EscapeDataString(destinations) + "&msg=" + Uri.EscapeDataString(message);

                HttpResponseMessage response = await client.GetAsync(ws_str);

                response.EnsureSuccessStatusCode();
                string responseBody = "";
                using (HttpContent content = response.Content)
                {
                    responseBody = await response.Content.ReadAsStringAsync();
                    //Console.WriteLine(responseBody + "........");
                }
                Library.WriteErrorLog(responseBody);
                return responseBody;
            }
        }


        public void SendEmail(string pTo, string pCc, string pBcc, string pSubject, string pBody, List<string> pAttachments)
        {
            using (var client = new HttpClient())
            {
                try
                {
                   // return "";
                    var portNumber = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["SendEmailPortNo"]);
                    var enableSSL = Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings["SendEmailEnableSSL"]);
                    var smtpAddress = Convert.ToString(ConfigurationManager.AppSettings["SendEmailSMTP"]);

                    var FromMailAddress = System.Configuration.ConfigurationManager.AppSettings["SendEmailFrom"].ToString();
                    var password = System.Configuration.ConfigurationManager.AppSettings["SendEmailFromPassword"].ToString();

                    //SmtpClient _client = new SmtpClient(ConfigurationManager.AppSettings["SMTPServer"]);
                    var _client = new SmtpClient(smtpAddress, portNumber) //Port 8025, 587 and 25 can also be used.
                    {
                        Credentials = new NetworkCredential(FromMailAddress, password),
                    };
                    _client.UseDefaultCredentials = false;
                    MailMessage _mailMessage = new MailMessage();
                    _mailMessage.To.Add(new MailAddress(pTo));
                    _mailMessage.From = new MailAddress(FromMailAddress, "GeneInsure");
                    _mailMessage.Subject = pSubject;
                    _mailMessage.IsBodyHtml = true;

                    //if (pAttachments != null)
                    //{
                    //    if (pAttachments[0] != "")
                    //    {
                    //        foreach (var item in pAttachments)
                    //        {
                    //            var fileinfo = new FileInfo();

                    //            System.Net.Mail.Attachment attachment;
                    //            attachment = new System.Net.Mail.Attachment(System.Web.HttpContext.Current.Server.MapPath(item.ToString()));
                    //            _mailMessage.Attachments.Add(attachment);
                    //        }

                    //    }

                    //}


                    AlternateView plainView = AlternateView.CreateAlternateViewFromString(pBody, null, "text/plain");
                    AlternateView htmlView = AlternateView.CreateAlternateViewFromString(pBody, null, "text/html");
                    _mailMessage.AlternateViews.Add(plainView);
                    _mailMessage.AlternateViews.Add(htmlView);
                    using (SmtpClient smtp = new SmtpClient(smtpAddress, portNumber))
                    {
                        smtp.Credentials = new NetworkCredential(FromMailAddress, password);
                        smtp.EnableSsl = enableSSL;
                        try
                        {
                            smtp.Send(_mailMessage);


                            Library.WriteErrorLog("EMAIL send succesful");
                        }
                        catch (Exception ex)
                        {
                            Library.WriteErrorLog("EMAIL send unsuccesful");
                        }
                    }
                    //populateMailAddresses(pTo, _message.To);

                    //if (pCc != null && pCc != "")
                    //    populateMailAddresses(pCc, _message.CC);
                    //if (pBcc != null && pBcc != "")
                    //    populateMailAddresses(pBcc, _message.Bcc);
                    //_message.Body = pBody;
                    //_message.BodyEncoding = System.Text.Encoding.UTF8;
                    //_message.Subject = pSubject;
                    //_message.SubjectEncoding = System.Text.Encoding.UTF8;
                    //_message.IsBodyHtml = pIsHTML;
                    //if (pAttachments != null)
                    //{
                    //    foreach (string _str in pAttachments)
                    //        _message.Attachments.Add(new Attachment(_str));
                    //}
                    //client.Send(_message);
                    //_message.Dispose();
                }
                catch (Exception ex)
                {


                    //WriteLog(ex.Message);

                }
            }

        }

    }
}
