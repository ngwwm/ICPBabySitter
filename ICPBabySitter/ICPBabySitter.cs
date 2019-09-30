using System;
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

namespace ICPBabySitter
{
    class ICPUser
    {
        private string m_userid;
        private string m_email;
        private string m_link;
        private bool m_found;
        //private readonly string m_filtersubject = "been invited to join HA Innovation Collaboration Platform";

        //constructor
        public ICPUser(EmailMessage email)
        {            
            m_found = ExtractRegistrationLink(email);
        }

        public bool IsFound()
        {
            return m_found;
        }

        private bool ExtractRegistrationLink(EmailMessage email)
        {
            string body;

            if (ConfigurationManager.AppSettings["debug_mode"] == "Y")
                Console.WriteLine(DateTime.Now.ToString() + " DEBUG Email Subject ==> " + email.Subject.ToString());

            //if (email.Subject.IndexOf(m_filtersubject) > 0)
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
                try
                {
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

                        if (screenEle is null || emailEle is null)
                        {
                            Console.WriteLine(DateTime.Now.ToString() + " INFO Fail to extract the UserId and Email from " + link);
                            return false;
                        }
                        if (ConfigurationManager.AppSettings["debug_mode"] == "Y")
                            Console.WriteLine(DateTime.Now.ToString() + " DEBUG Extracted values ==> " + screenEle.value + "," + emailEle.value + "," + link);

                        m_userid = screenEle.value;
                        m_email = emailEle.value;
                        m_link = link;
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(DateTime.Now.ToString() + " ERROR " + ex.Message);
                    return false;
                }
            }
            //return false;
        }

        public bool Save2DB(SqlConnection objConn)
        {
            try
            {
                SqlCommand objCommand = new SqlCommand("create_user", objConn);
                objCommand.CommandType = CommandType.StoredProcedure;
                objCommand.CommandTimeout = 60;

                objCommand.Parameters.Add("@userid", SqlDbType.VarChar, 20).Value = m_userid;
                objCommand.Parameters.Add("@email", SqlDbType.VarChar, 50).Value = m_email;
                objCommand.Parameters.Add("@link", SqlDbType.VarChar, 150).Value = m_link;

                //inboxFolder.Items[i].GetInspector.Close(Outlook.OlInspectorClose.olDiscard);

                if ((int)objCommand.ExecuteScalar() != 1)
                {
                    Console.WriteLine(DateTime.Now.ToString() + " ERROR Error saving record: " + m_userid + ".");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(DateTime.Now.ToString() + " ERROR " + ex.Message);
                return false;
            }
            return true;
        }

    }
 
    class ICPBabySitter
    {
        static ExchangeService service;

        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                PrintUsage();
                return;
            }

            foreach (string arg in args)
            {
                switch (arg.Substring(0, 2).ToUpper())
                {
                    case "-E":
                        // Check fax for receipt
                        SetupEWS();
                        ExtractLinks();
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

        private static void SetupEWS()
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
            /*
                        SecureString lanPwd = new SecureString();
                        lanPwd.AppendChar('p');
                        lanPwd.AppendChar('w');
                        lanPwd.AppendChar('d');
            */
            //SecureString lanPwd = new SecureString("password", 8);
            service.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["ews_user"], ConfigurationManager.AppSettings["ews_pwd"], "CORP");

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
        
        private static int ExtractLinks()
        {
            //PostItem postItem;
            EmailMessage emailMessage;
            //String emailEntityId = "";
            int pageSize, offset = 0;
            int total = 0;
            SqlConnection objConn = null;

            //Open DB connection
            try
            {
                objConn = new SqlConnection(ConfigurationManager.ConnectionStrings["ICPConn"].ConnectionString);
                objConn.Open();
            }
            catch (Exception ex)
            {
                Console.WriteLine(DateTime.Now.ToString() + " ERROR " + ex.Message);
                return -1;
            }
            //Console.WriteLine("Limited to process  " + findResults.Items.Count.ToString() + " email(s).");
            pageSize = Convert.ToInt32(ConfigurationManager.AppSettings["num_of_email"]);

            //Bind to Inbox folder
            Folder inboxfolder = Folder.Bind(service, WellKnownFolderName.Inbox);

            //Retrieve first n emails items
            FindItemsResults<Item> findResults;

            SearchFilter srchfiltercoll = new SearchFilter.SearchFilterCollection(LogicalOperator.And, 
                new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false), 
                new SearchFilter.ContainsSubstring(ItemSchema.Subject, "been invited to join HA Innovation Collaboration Platform"));

            //SearchFilter srchfilter = new SearchFilter.ContainsSubstring(ItemSchema.Subject, "been invited to join HA Innovation Collaboration Platform");

            do
            {
                findResults = service.FindItems(WellKnownFolderName.Inbox, srchfiltercoll, new ItemView(pageSize, offset));

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
                            if (icpuser.IsFound())
                            {
                                if (icpuser.Save2DB(objConn))
                                {
                                    //mark the message as read
                                    //emailMessage.IsRead = true;
                                    //item.Update(ConflictResolutionMode.AutoResolve);
                                }
                            }
                        }
                        total += 1;
                    }
                }
                Console.WriteLine(DateTime.Now.ToString() + " INFO  " + total + " emails were processed so far.");
                offset += pageSize;
            } while (findResults.MoreAvailable);

            if (objConn.State == ConnectionState.Open)
                objConn.Close();

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
            Console.WriteLine("Usage: Required either [-E]");
            Console.WriteLine("  -E: Extract Links");
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
