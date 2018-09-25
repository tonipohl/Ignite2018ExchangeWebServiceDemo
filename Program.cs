using System;
using System.Configuration;
using System.Linq;
using System.Net;
using Microsoft.Exchange.WebServices.Data;
using System.Globalization;

// Demo for Ignite 2018: Modernize your apps with Graph
// Toni Pohl, atwork.at, @atwork
// Legacy App: Send an email to a resource mailbox with EWS and create an appointment out of it.
//
// Subject: must contain "ID:"
//      ID:45 Modern Workplace Conference 
// Body: must contain the startdate and enddate as 1st and 2nd line as here:
//      20180902 09:00
//      20180902 15:00
// To delete an existing ID, send "delete ID:45" as subject.
//
namespace ClassesTool
{
    class Program
    {
        private static ExchangeService service;

        static void Main(string[] args)
        {
            Console.WriteLine("ClassesTool-EWS");
            var mailboxes = ConfigurationManager.AppSettings["Inbox"].Split(new[] { ';' }).ToList();
            var username = ConfigurationManager.AppSettings["User"];
            var password = Encrypt.DecryptString(ConfigurationManager.AppSettings["SecuredPassword"]);

            // Connect to EXO
            Console.WriteLine("Connect to Exchange Online");
            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
            service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Credentials = new WebCredentials(username, password);
            service.AutodiscoverUrl(username, RedirectionUrlValidationCallback);

            // Do something
            Console.WriteLine("Loop through mailboxes");
            foreach (var box in mailboxes)
            {
                //SendEmail(box);
                GetEmails(box);
                //DeleteAllAppointments(box);
            }

            Console.WriteLine("\n-- Done. Press any key to end. --");
            Console.ReadKey();
        }

        private static void MainWithComments()
        {
            var username = ConfigurationManager.AppSettings["User"];
            var password = ConfigurationManager.AppSettings["Password"];

            // See more about EWS:
            // https://docs.microsoft.com/en-us/previous-versions/office/developer/exchange-server-2010/jj220499(v%3dexchg.80)

            // Add the certificate validation callback method to the ServicePointManager.
            // This callback has to be available before you make any calls to the EWS end point.
            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;

            // Instantiate the ExchangeService object with the service version you intend to target.
            service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);

            // For Exchange Online service, pass explicit credentials. 
            service.Credentials = new WebCredentials(username, password);
            // In domain-joined, we could use the user's Windows credentials instead:
            // service.UseDefaultCredentials = true;

            // The client is ready to make the first call to the Autodiscover service 
            // to get the service URL for calls to the EWS service.

            // The AutodiscoverUrl method on the ExchangeService object performs a call 
            // to the Autodiscover service to get the service URL. If this call is successful, 
            // the URL property on the ExchangeService object will be set with the service URL. 
            // Pass the user principal name and the RedirectionUrlValidationCallback to the AutodiscoverUrl method.

            // If you want to see the actual calls being made enable Tracing.
            //service.TraceEnabled = true;
            //service.TraceFlags = TraceFlags.All;

            service.AutodiscoverUrl(username, RedirectionUrlValidationCallback);
        }

        private static bool CertificateValidationCallBack(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate, System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
            {
                return true;
            }

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                           (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            continue;
                        }
                        else
                        {
                            if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                return false;
                            }
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are 
                // untrusted root errors for self-signed certificates. These certificates are valid
                // for default Exchange server installations, so return true.
                return true;
            }
            else
            {
                // In all other cases, return false.
                return false;
            }
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        private static void SendEmail(string mailbox)
        {
            // Generate n demo emails
            for (int i = 0; i < 3; i++)
            {
                Random rnd = new Random();
                int randomID = rnd.Next(10, 100);   // Generate a random ID 10..99 for the new appointment

                EmailMessage email = new EmailMessage(service);

                // The sender is the user we used in the config file...
                email.ToRecipients.Add(mailbox);
                email.Subject = "ID:" + randomID.ToString() + " Workshop";
                email.Body = new MessageBody(DateTime.Now.AddHours(randomID).ToString("yyyyMMdd HH:mm") + Environment.NewLine + "<br>" +
                    DateTime.Now.AddHours(randomID + 2).ToString("yyyyMMdd HH:mm") + Environment.NewLine + "<br>" +
                    "This is a generated seminar entry with ID:" + randomID.ToString() + Environment.NewLine + " sent by the SeminarCalendarTool.\n\r");
                email.SendAndSaveCopy();
                Console.WriteLine("Sent: {0}", email.Subject);
            }
        }

        private static Folder GetMailboxInbox(string mailbox)
        {
            try
            {
                var box = new Mailbox(mailbox);
                var rootfolder = Folder.Bind(service, new FolderId(WellKnownFolderName.Inbox, box));
                Console.WriteLine("-- {0} --", box.Address);
                return rootfolder;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error GetMailbox: " + ex.Message);
                return null;
            }
        }

        private static void GetEmails(string mailbox)
        {
            // https://docs.microsoft.com/en-us/previous-versions/office/developer/exchange-server-2010/dd633640(v%3Dexchg.80)
            // https://code.msdn.microsoft.com/exchange/exchange-2013-101-code-3c38582c
            var messageCount = Convert.ToInt32(ConfigurationManager.AppSettings["MessageCount"]);
            Console.WriteLine("GetEmails: {0}", mailbox);

            var folder = GetMailboxInbox(mailbox);

            if (folder != null)
            {
                ItemView view = new ItemView(messageCount);
                foreach (var item in folder.FindItems(view))
                {
                    try
                    {
                        // Get the email object & don't forget to add the TextBody property to get it!
                        var msg = EmailMessage.Bind(service, item.Id,
                            new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.TextBody));

                        Console.WriteLine("E-Mail: {0}, {1}", msg.DateTimeSent, msg.Subject);
                        // valid appointment? does it contain "ID:" ?
                        if (msg.Subject.ToLower().Contains("id:"))
                        {
                            if (msg.Subject.ToLower().StartsWith("delete"))
                            {
                                DeleteCalendarEntry(mailbox, msg.Subject);
                            }
                            else
                            {
                                AddAppointment(mailbox, msg.Subject, msg.TextBody);
                            }
                            // remove that email to..somewhere else
                            DeleteMail(mailbox, msg);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error GetMailbox: " + ex.Message);
                        continue;
                    }
                }
            }
        }

        private static void AddAppointment(string mailbox, string subject, string body)
        {
            DateTime startdate, enddate;
            body = body.Replace("\n", ""); // Remove linefeed
            string[] lines = body.Split('\r'); // split by line

            if (lines.Count() > 1)
            {
                try
                {
                    startdate = GetDate(lines[0].Trim());
                    enddate = GetDate(lines[1].Trim());
                    body = string.Format("Timestamp: {0}", DateTime.Now.ToString("G")) + "<br>" + body;
                    Console.WriteLine("Creating: {0} to {1}. {2}", startdate, enddate, subject);

                    CreateNewCalendarEntry(mailbox, startdate, enddate, subject, body);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error AddAppointment: " + ex.Message);
                }
            }
        }

        private static void CreateNewCalendarEntry(string mailbox, DateTime startdate, DateTime enddate, string subject, string body)
        {
            // create a new appointment
            var appointment = new Appointment(service);

            appointment.Start = startdate;
            appointment.End = enddate;
            appointment.Subject = subject;
            appointment.Body = body;
            appointment.StartTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
            appointment.EndTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");

            var box = new Mailbox(mailbox);
            var calendar = Folder.Bind(service, new FolderId(WellKnownFolderName.Calendar, box));
            appointment.Save(calendar.Id);
        }

        private static void DeleteCalendarEntry(string mailbox, string subject)
        {
            if (subject.ToLower().Contains("id:"))
            {
                try
                {
                    // search for the subject without the "delete" command
                    subject = subject.ToLower().Replace("delete", "").Trim();

                    var box = new Mailbox(mailbox);
                    var calendar = Folder.Bind(service, new FolderId(WellKnownFolderName.Calendar, box));

                    ItemView view = new ItemView(10);
                    // and search for the (part of) the subject: "id:10"
                    FindItemsResults<Item> results = calendar.FindItems(subject, view);

                    foreach (var oneappointment in results)
                    {
                        // Ensure it's an appointment (and not a meeting that cannot be deleted)
                        if (oneappointment is Appointment)
                        {
                            Console.WriteLine("Delete appointment: {0}", oneappointment.Subject);
                            oneappointment.Delete(DeleteMode.SoftDelete);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error DeleteCalendarEntry: " + ex.Message);
                }
            }
        }

        private static void DeleteMail(string mailbox, EmailMessage message)
        {
            if (ConfigurationManager.AppSettings["MoveEmails"].ToLower() == "yes")
            {
                // delete the message
                message.Delete(DeleteMode.MoveToDeletedItems);
                Console.WriteLine("Delete email: {0}", message.Subject);
            }
        }

        private static DateTime GetDate(string date)
        {
            // Just a helper to convert a string to a DateTime
            var appointmentDate = DateTime.Today;
            try
            {
                date = date.Replace("T", "");
                // "MM/dd/yyyy HH:mm"
                appointmentDate = DateTime.ParseExact(date, "yyyyMMdd HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error GetDate: " + ex.Message);
            }
            return appointmentDate;
        }

        private static void DeleteAllAppointments(string mailbox)
        {
            var box = new Mailbox(mailbox);
            var calendar = Folder.Bind(service, new FolderId(WellKnownFolderName.Calendar, box));

            ItemView view = new ItemView(100);
            foreach (var item in calendar.FindItems(view))
            {
                try
                {
                    // Get the appointment object
                    var message = Appointment.Bind(service, item.Id,
                        new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.Subject));

                    Console.WriteLine("Deleting: {0}, {1}, {2}", message.Start, message.End, message.Subject);
                    message.Delete(DeleteMode.HardDelete);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error DeleteAllAppointments: " + ex.Message);
                    continue;
                }
            }
        }
    }
}
// end of Program
