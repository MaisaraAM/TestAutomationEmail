using ExcelSol;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace TestAutomationEmail.Pages
{
    public class EmailPage : TestFixtureBase
    {
        static Outlook.Application application = new Outlook.Application();
        static Outlook.Accounts accounts = application.Session.Accounts;

        public string searchEmail(string subjects)
        {
            string HTMLBody = null;

            foreach (Outlook.Account account in accounts)
            {
                if (string.Equals(account.DisplayName.Trim(), "t-mamaher@EFG-HERMES.com", StringComparison.InvariantCultureIgnoreCase))
                {
                    Console.WriteLine("Current email account is: " + account.DisplayName);

                    Outlook.Application oApp = new Outlook.Application();
                    Outlook.NameSpace oNS = oApp.GetNamespace("mapi");

                    oNS.SendAndReceive(false);
                    Thread.Sleep(5000);

                    oNS.Logon(Missing.Value, Missing.Value, false, true);

                    Outlook.MAPIFolder oInbox = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                    Outlook.Items oItems = oInbox.Items;

                    oItems.Sort("[ReceivedTime]", false);

                    Outlook.MailItem oMsg = (Outlook.MailItem)oItems[oItems.Count];

                    string filename = "";
                    bool emailFound = false;
                    int itr = 0;

                    while (emailFound == false)
                    {
                        for (int i = oItems.Count; i >= 1; i--)
                        {
                            try
                            {
                                oMsg = (Outlook.MailItem)oItems[i];

                                //Output some common properties.
                                string subj = oMsg.Subject;
                                string sender = oMsg.SenderName;
                                string reciTime = oMsg.ReceivedTime.ToString();
                                string Body = oMsg.Body;
                                HTMLBody = oMsg.HTMLBody;

                                DateTime reciTimeDT = DateTime.Parse(reciTime);

                                if (!String.IsNullOrEmpty(subj))
                                {
                                    if (subj == subjects.Trim())
                                    {
                                        emailFound = true;
                                        sender = oMsg.SenderName;
                                        reciTime = oMsg.ReceivedTime.ToString();
                                        Body = oMsg.Body;
                                        HTMLBody = oMsg.HTMLBody;

                                        break;
                                    }
                                }
                            }

                            catch { }
                        }

                        if (emailFound == false)
                        {
                            oNS.SendAndReceive(false);
                            Thread.Sleep(5000);
                        }

                        itr += 1;

                        if (itr == 3)
                            break;
                    }

                    Thread.Sleep(2000);
                }
            }

            return HTMLBody;
        }

    }
}
