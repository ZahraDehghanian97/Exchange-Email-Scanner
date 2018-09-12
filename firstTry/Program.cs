using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Security;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Exchange.WebServices.Data;
namespace firstTry
{
    class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("No Auto sign in is discoverd");
            /* Console.WriteLine("enter your email account");
             string emailS = Console.ReadLine();
             Console.WriteLine("enter your password key");
             string pass = Console.ReadLine();
 */         System.Console.WriteLine("try to connect to exchange server ...");
            string emailS = "m@z.d";
            string pass = "1qaz!QAZ";
            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Credentials = new WebCredentials(emailS, pass);
            service.AutodiscoverUrl(emailS, RedirectionUrlValidationCallback);
            Console.WriteLine("sign in");

            string checkSample = System.IO.File.ReadAllText("C:\\temp\\checkSample.txt");
            System.Console.WriteLine("Contents of sample file for checking : " + checkSample);
            int deletedFiles = 0;
            ItemView view = new ItemView(1);
            string querystring = " Kind:email"; //HasAttachments:true
            FindItemsResults<Item> results = service.FindItems(WellKnownFolderName.Inbox, querystring, view);
            System.Console.WriteLine("number of email in your inbox: " + results.TotalCount);
            if (results.TotalCount > 0)
            {
                foreach (Item item in results)
                {
                    EmailMessage email = item as EmailMessage;
                    //System.Console.WriteLine(email.TextBody);
                    email.Load(new PropertySet(EmailMessageSchema.Attachments));
                    //System.Console.WriteLine("body of email: " + email.TextBody);
                    foreach (Attachment attachment in email.Attachments)
                    {

                        if (attachment is FileAttachment)
                        {
                            FileAttachment fileAttachment = attachment as FileAttachment;
                            fileAttachment.Load();
                            Console.WriteLine("Load attachment with a name = " + fileAttachment.Name);
                            fileAttachment.Load("C:\\temp\\" + fileAttachment.Name);
                            string text = System.IO.File.ReadAllText("C:\\temp\\" + fileAttachment.Name);
                            //System.Console.WriteLine("Contents of WriteText.txt = {0}", text);
                            if (text.Contains(checkSample))
                            {
                                System.Console.WriteLine("this file has to be deleted ");
                                email.Attachments.Remove(fileAttachment);
                                email.Update(ConflictResolutionMode.AlwaysOverwrite);
                                deletedFiles++;
                                if (email.Attachments.Count == 0) {break; }
                            }
                            else
                            {
                                System.Console.WriteLine("there is no problem with this file");
                            }
                        }

                    }
                }
            }
            Console.WriteLine("checking is completed");
            Console.WriteLine("find " + deletedFiles + " threat . delete all bad file(s)");
            Console.WriteLine("insert enter key to continue ... ");
            Console.ReadLine();

        }
        /* private static void ProcessItem(Item item,int count) {
             Console.WriteLine("processing email number "+count );

         }
         */
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

    }
}
