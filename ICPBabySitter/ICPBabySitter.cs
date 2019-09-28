﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Exchange.WebServices.Data;
using mshtml;
using System.Security;

namespace ICPExtractRegLinks
{
    class ICPUser
    {
        private string m_userid;
        private string m_email;
        private string m_link;
        private int m_valid;
        private EmailMessage m_emailmsg;

        public ICPUser(EmailMessage email)
        {
            m_emailmsg = email;
            m_valid = ExtractRegistrationLink(email);
        }

        public Boolean IsValid()
        {
            return (m_valid == 1);
        }

        private int ExtractRegistrationLink(EmailMessage email)
        {
            string body;            

            if (ConfigurationManager.AppSettings["debug_mode"] == "Y")
                Console.WriteLine("DEBUG  " + email.Subject.ToString());

            if (email.Subject.IndexOf("been invited to join HA Innovation Collaboration Platform") > 0)
            {
                body = email.Body;                

                int linkStart = body.LastIndexOf("https://");
                int linkEnd = body.LastIndexOf("</a>");

                string link = body.Substring(linkStart, linkEnd - linkStart);

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(link);
                request.Method = "GET";
                /*
                                IWebProxy proxy = request.Proxy;
                                // Print the Proxy Url to the console.
                                if (proxy != null)
                                {
                                    Console.WriteLine("Proxy: {0}", proxy.GetProxy(request.RequestUri));
                                }
                                else
                                {
                                    Console.WriteLine("Proxy is null; no proxy will be used");
                                }
                */
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                using (System.IO.StreamReader reader = new System.IO.StreamReader(response.GetResponseStream()))
                {
                    string sPage = reader.ReadToEnd();
                    // Process the response text if you need to...                                        
                    object[] oPageText = { sPage };

                    HTMLDocument doc = new HTMLDocument();
                    IHTMLDocument2 doc2 = (IHTMLDocument2)doc;
                    doc2.write(oPageText);

                    HTMLInputElement screenEle = (HTMLInputElement)doc.getElementById("b_register_screen_name");
                    HTMLInputElement emailEle = (HTMLInputElement)doc.getElementById("b_register_email");

                    Console.WriteLine(screenEle.value + "," + emailEle.value + "," + link);
                    m_userid = screenEle.value;
                    m_email = emailEle.value;
                    m_link = link;                                       
                    return 1;
                }
            }
            return 0;
        }

        public void CreateDBRecord(SqlConnection objConn)
        {
            SqlCommand objCommand = new SqlCommand("create_user", objConn);
            objCommand.CommandType = CommandType.StoredProcedure;
            objCommand.CommandTimeout = 60;

            objCommand.Parameters.Add("@userid", SqlDbType.VarChar, 20).Value = m_userid;
            objCommand.Parameters.Add("@email", SqlDbType.VarChar, 50).Value = m_email;
            objCommand.Parameters.Add("@link", SqlDbType.VarChar, 150).Value = m_link;

            //inboxFolder.Items[i].GetInspector.Close(Outlook.OlInspectorClose.olDiscard);

            try
            {

                objConn.Open();

                int numberOfRecords = (int)objCommand.ExecuteScalar();

                objConn.Close();

                if (numberOfRecords == 1)
                {
                    if (ConfigurationManager.AppSettings["debug_mode"] == "N")
                        m_emailmsg.IsRead = true;
                    item.Update(ConflictResolutionMode.AutoResolve);
                    Console.WriteLine(DateTime.Now.ToString() + ": Fax Subject - " + item.Subject + " SENT.");
                }

            }
            catch (Exception ex)
            {
                objConn.Close();
                Console.WriteLine(DateTime.Now.ToString() + ": CheckFax failed.");
                Console.WriteLine(ex.Message);
            }
        }

    }

    class UserLink
    {
        private string m_userid;
        private string m_email;
        private string m_link;

        private void SetUserLink(string userid, string email, string link)
        {
            m_userid = userid;
            m_email = email;
            m_link = link;
        }

    }
    
    class ICPBabySitter
    {
        static ExchangeService service;

        private static string m_userid;
        private static string m_email;
        private static string m_link;

        private static void SetUser(string userid, string email, string link)
        {
            m_userid = userid;
            m_email = email;
            m_link = link;
        }

        static void Main(string[] args)
        {
            int start = 0;

            if (args.Length == 0)
            {
                PrintUsage();
                return;
            }

            foreach (string arg in args)
            {
                switch (arg.Substring(0, 2).ToUpper())
                {
                    case "-C":
                        // Check fax for receipt
                        setupEWS();
                        CheckFax();
                        break;
                    /*
                    case "-S":
                        // Send fax in MAIL_DB
                        setupEWS();
                        SendFax();
                        break;
                    */
                    case "-?":
                        // Print usage
                        PrintUsage();
                        break;
                    default:
                        // Print usage
                        PrintUsage();
                        break;
                }
            }
        }

        private static void setupEWS()
        {
            //Add your certificate validation callback method to the ServicePointManager by adding 
            //the following code to the beginning of the Main(string[] args) method. 
            //This callback has to be available before you make any calls to the EWS end point.
            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;

            //instantiate the ExchangeService object with the service version you intend to target.
            service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);

            //Domain-joined clients that target an on-premise Exchange server can use the default credentials 
            //of the user who is logged on, assuming the credentials are associated with a mailbox.
            //service.UseDefaultCredentials = true;

            //service.Credentials = new NetworkCredential("username", "password", "CORP"); ;
            string landId = ConfigurationManager.AppSettings["client_email"];
            string lanPwd = ConfigurationManager.AppSettings["client_key"];
            /*
                        SecureString lanPwd = new SecureString();
                        lanPwd.AppendChar('p');
                        lanPwd.AppendChar('w');
                        lanPwd.AppendChar('d');
            */
            //SecureString lanPwd = new SecureString("password", 8);
            service.Credentials = new NetworkCredential(landId, lanPwd, "CORP");

            //The AutodiscoverUrl method on the ExchangeService object performs a call to the Autodiscover service 
            //to get the service URL. If this call is successful, the URL property on the ExchangeService object 
            //will be set with the service URL. Pass the user principal name and the 
            //RedirectionUrlValidationCallback to the AutodiscoverUrl method.
            service.AutodiscoverUrl(ConfigurationManager.AppSettings["client_email"], RedirectionUrlValidationCallback);

            //service.Url = new Uri("https://mailcorphc02.corp.ha.org.hk/EWS/Exchange.asmx");
            //service.Url = new Uri(ConfigurationManager.AppSettings["ews_auto_discover"]);
            if (ConfigurationManager.AppSettings["debug_mode"] == "Y")
                Console.WriteLine(service.Url);
            //service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, "hasdevsa@ho.ha.org.hk");

        }

        //execute every 3 minutes
        private static int CheckFax()
        {
            PostItem postItem;
            EmailMessage emailMessage;
            String emailEntityId = "";
            int offset = 0;

            //Bind to Inbox folder
            Folder inboxfolder = Folder.Bind(service, WellKnownFolderName.Inbox);

            //Console.WriteLine("The " + inboxfolder.DisplayName + " has " + inboxfolder.ChildFolderCount + " child folders.");
            //Console.WriteLine("The " + inboxfolder.DisplayName + " has " + inboxfolder.TotalCount + " items.");
            
            //Retrieve first xx emails items
            FindItemsResults<Item> findResults = service.FindItems(
               WellKnownFolderName.Inbox,
               new ItemView(Convert.ToInt32(ConfigurationManager.AppSettings["num_of_email"]), offset));


            Console.WriteLine("Limited to process  " + findResults.Items.Count.ToString() + " email(s).");
            //service.LoadPropertiesForItems(findResults.Items, PropertySet.FirstClassProperties);
            foreach (Item item in findResults.Items)
            {
                if (item.ItemClass == "IPM.Note")
                {                    
                    emailMessage = (EmailMessage)item;
                    //Console.WriteLine("IsRead: " + emailMessage.IsRead);

                    //perform other actions here...

                    if (!emailMessage.IsRead)
                    {                        
                        //item.Load();
                        //emailEntityId = ConvertEWSidToEntryID(service, item.Id.ToString(), ConfigurationManager.AppSettings["client_email"]);                        

                        emailMessage.Load();

                        ICPUser icpuser = new ICPUser(emailMessage);

                        //if (ExtractRegistrationLink(emailMessage) > 0) {
                        if (icpuser.IsValid())
                        {
                            using (SqlConnection objConn = new SqlConnection(ConfigurationManager.ConnectionStrings["ICPConn"].ConnectionString))
                            {
                                SqlCommand objCommand = new SqlCommand("create_user", objConn);
                                objCommand.CommandType = CommandType.StoredProcedure;
                                objCommand.CommandTimeout = 60;

                                objCommand.Parameters.Add("@userid", SqlDbType.VarChar, 20).Value = m_userid;
                                objCommand.Parameters.Add("@email", SqlDbType.VarChar, 50).Value = m_email;
                                objCommand.Parameters.Add("@link", SqlDbType.VarChar, 150).Value = m_link;

                                //inboxFolder.Items[i].GetInspector.Close(Outlook.OlInspectorClose.olDiscard);


                                try
                                {

                                    objConn.Open();

                                    int numberOfRecords = (int)objCommand.ExecuteScalar();

                                    objConn.Close();

                                    if (numberOfRecords == 1)
                                    {
                                        if (ConfigurationManager.AppSettings["debug_mode"] == "N")
                                            emailMessage.IsRead = true;
                                        item.Update(ConflictResolutionMode.AutoResolve);
                                        Console.WriteLine(DateTime.Now.ToString() + ": Fax Subject - " + item.Subject + " SENT.");
                                    }

                                }
                                catch (Exception ex)
                                {                                    
                                    objConn.Close();
                                    Console.WriteLine(DateTime.Now.ToString() + ": CheckFax failed.");
                                    Console.WriteLine(ex.Message);
                                }
                            }
                        }
                        
                        //debug
                        /*
                        if (ConfigurationManager.AppSettings["debug_mode"] == "Y")
                        {
                            Console.WriteLine("Sender: " + emailMessage.Sender.Name);
                            Console.WriteLine("Subject: " + item.Subject);
                            Console.WriteLine("Id: " + item.Id);
                            Console.WriteLine("ItemClass: " + item.ItemClass);
                            Console.WriteLine("EntityId: " + emailEntityId);
                            Console.WriteLine("IsRead: " + emailMessage.IsRead);                            
                        }
                        */
                    }
                }
                //for other ItemClass
                else
                {
                    //perform other actions here...
                }
            }
            return 0;
        }

        private static int XXXExtractRegistrationLink(EmailMessage email)
        {
            string body;

            if (ConfigurationManager.AppSettings["debug_mode"] == "Y")
                Console.WriteLine("DEBUG  " + email.Subject.ToString());

            if (email.Subject.IndexOf("been invited to join HA Innovation Collaboration Platform") > 0)
            {
                body = email.Body;
                //Console.WriteLine(body);
                
                int linkStart = body.LastIndexOf("https://");
                int linkEnd = body.LastIndexOf("</a>");

                string link = body.Substring(linkStart, linkEnd - linkStart);

  //              Console.WriteLine(link);
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(link);
                request.Method = "GET";
/*
                IWebProxy proxy = request.Proxy;
                // Print the Proxy Url to the console.
                if (proxy != null)
                {
                    Console.WriteLine("Proxy: {0}", proxy.GetProxy(request.RequestUri));
                }
                else
                {
                    Console.WriteLine("Proxy is null; no proxy will be used");
                }
*/
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                using (System.IO.StreamReader reader = new System.IO.StreamReader(response.GetResponseStream()))
                {
                    string sPage = reader.ReadToEnd();
                    // Process the response text if you need to...                                        
                    object[] oPageText = { sPage };

                    HTMLDocument doc = new HTMLDocument();
                    IHTMLDocument2 doc2 = (IHTMLDocument2) doc;
                    doc2.write(oPageText);

                    HTMLInputElement screenEle = (HTMLInputElement) doc.getElementById("b_register_screen_name");
                    HTMLInputElement emailEle = (HTMLInputElement) doc.getElementById("b_register_email");
                    
                    Console.WriteLine(screenEle.value + "," + emailEle.value + "," + link);
                    SetUser(screenEle.value, emailEle.value, link);
                    return 1;
                    //b_register_screen_name
                    //b_register_email            
                    //Console.WriteLine(doc.getElementById("b_register_screen_name"));
                    //Console.WriteLine(doc.getElementById("b_register_email"));
                }                
            }
            return 0;
        }

        private static string ConvertEWSidToEntryID(ExchangeService service, string idToConvert, string mailboxSMTP)
        {
            // Specify the item or folder identifier, the identifier type, the SMTP address of the mailbox
            // that contains the identifier, and whether the item/folder identifier represents an archived
            // item or folder.
            AlternateId originalId = new AlternateId(IdFormat.EwsId, idToConvert, mailboxSMTP, false);

            // Send a request to convert the item identifier. This results in a call to EWS.
            AlternateId newId = service.ConvertId(originalId, IdFormat.HexEntryId) as AlternateId;

            Console.WriteLine("Original identifier: " + idToConvert);
            Console.WriteLine("Converted identifier type: " + newId.Format);
            Console.WriteLine("Converted identifier: " + newId.UniqueId);

            return newId.UniqueId;
        }

        private static void PrintUsage()
        {
            Console.WriteLine("Usage: Required either [-C]");
            Console.WriteLine("  -C: Check fax in Inbox");
            //Console.WriteLine("  -S: Send fax in MAIL_DB");
        }

        //Create a certificate validation callback method
        private static bool CertificateValidationCallBack(
            object sender,
            System.Security.Cryptography.X509Certificates.X509Certificate certificate,
            System.Security.Cryptography.X509Certificates.X509Chain chain,
            System.Net.Security.SslPolicyErrors sslPolicyErrors)
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

        //validates whether the redirected URL is using Transport Layer Security.
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
    }
}