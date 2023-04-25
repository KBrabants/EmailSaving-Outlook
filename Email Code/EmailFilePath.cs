using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailSaving
{
    internal class EmailFilePath
    {
        public static Accounts LoggedInOutlookAccounts()
        {

            Microsoft.Office.Interop.Outlook.Application outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
            NameSpace outlookNamespace = outlookApplication.GetNamespace("MAPI");
            return outlookNamespace.Accounts;

        }
        public static List<string> GetFilePath(string customerEmail, GeoAccounts Geoaccounts, bool SaveIfNoAccountNumber = false, bool CheckName = false)
        {
            List<string> AccountsFound = new List<string>();

            //string mkkFolders = "S:\\Microkey Subscriber Documents\\";

            string mkkFolders = "S:\\Microkey Subscriber Documents\\";

            string EmailNotAttachedToAccountPath = "S:\\OutLook Emails Saved\\";

            int locationsFound = 0;

            if (customerEmail == null)
            {
                return new List<string>();
            }

            customerEmail = customerEmail.Replace("'", "");
            customerEmail = customerEmail.Replace(" ", "");
            if (customerEmail.ToUpper().Contains("@ALARMNET.COM") || customerEmail == "")
            {
                AccountsFound.Add(EmailNotAttachedToAccountPath + "General" + "\\");
                return AccountsFound;
            }

            if (Geoaccounts.list != null)
            {
                foreach (GeoAccount account in Geoaccounts.list)
                {

                    if (account.Email == null || customerEmail == null)
                    {
                        continue;
                    }
                    string temp = account.Email.Replace("'", "");
                    temp = temp.Replace(" ", "");
                    temp = temp.ToUpper();

                    string temp2 = customerEmail.ToUpper();
                    temp2 = temp2.Replace(" ", "");
                    temp2 = temp2.Replace("'", "");


                    if (temp.ToUpper() == temp2)
                    {
                        locationsFound++;
                        AccountsFound.Add(mkkFolders + account.SubId + "\\Automated Email Saves\\");
                        //Console.WriteLine(mkkFolders + account.SubId);
                    }

                }


            }

            if (SaveIfNoAccountNumber && locationsFound == 0)
            {
                AccountsFound.Add(EmailNotAttachedToAccountPath + customerEmail + "\\");
            }

            return AccountsFound;
        }
    }
}
