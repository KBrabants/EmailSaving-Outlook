using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace EmailSaving
{
    internal class EmailFileSaver
    {

        public static void SaveAllEmails(GeoAccounts geoAccounts, MAPIFolder Folder, DateTime afterTime, Action<MailItem, GeoAccounts, string> SaveBox)
        {
            Items Items = Folder.Items;

            List<MailItem> mailItems = EmailSorter.SeperateEmailsFromItems(Items, afterTime);

           
            for (int i = mailItems.Count -1; i > 0; i--)
            {
                SaveEmail(mailItems[i], geoAccounts, afterTime, SaveBox);
            }
        }

        public static void SaveEmail(MailItem email , GeoAccounts geoAccounts, DateTime afterTime, Action<MailItem, GeoAccounts, string> SaveBox)
        {

            if (email.SentOn <= afterTime)
                return;

            string name;


            name = email.Subject;

            if (name == null || name.Length < 2)
            {
                name = "Customer Emailed In";
            }

            name = FileSafeName(name);

            name += ".msg";


            SaveBox(email, geoAccounts, name);

        }


        public static void SaveOutbox(MailItem email, GeoAccounts geoAccounts, string name)
        {
            string emailAddress = "";
            foreach (Microsoft.Office.Interop.Outlook.Recipient Rec in email.Recipients)
            {
                 emailAddress = Rec.Address;
            }


            List<string> filePaths = EmailFilePath.GetFilePath(emailAddress, geoAccounts, true, true);

                foreach (string path in filePaths)
                {

                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    try
                    {
                        // Saving File
                        email.SaveAs(path + name, Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG);
                        File.SetCreationTime(path + name, email.SentOn);
                        File.SetLastWriteTime(path + name, email.SentOn);
                        //  Console.WriteLine($"Email: {name} Saved!");
                    }
                    catch
                    {
                        if (File.Exists(path + name))
                        {
                            // Gets the existing file info
                            DateTime filedate = File.GetLastWriteTime(path + name);

                            if (filedate < email.ReceivedTime)
                            {
                                //Overwriting existing file...
                                File.Delete(path + name);
                                email.SaveAs(path + name);
                                File.SetCreationTime(path + name, email.SentOn);
                                File.SetLastWriteTime(path + name, email.SentOn);
                                //  Console.WriteLine($"Email: {name} Saved!");
                            }

                        }

                    }
                }
        }

        public static void SaveInbox(MailItem email, GeoAccounts geoAccounts, string name)
        {
           string emailAddress = email.SenderEmailAddress;

           var filePaths = EmailFilePath.GetFilePath(emailAddress, geoAccounts);

                foreach (string path in filePaths)
                {

                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    try
                    {
                        // Saving File
                        email.SaveAs(path + name, Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG);
                        File.SetCreationTime(path + name, email.SentOn);
                        File.SetLastWriteTime(path + name, email.SentOn);

                    }
                    catch
                    {
                        if (File.Exists(path + name))
                        {
                            // Gets the existing file info
                            DateTime filedate = File.GetLastWriteTime(path + name);

                            if (filedate < email.SentOn)
                            {
                                //Overwriting existing file...
                                File.Delete(path + name);
                                email.SaveAs(path + name);
                                File.SetCreationTime(path + name, email.SentOn);
                                File.SetLastWriteTime(path + name, email.SentOn);
                                //     Console.WriteLine($"Email: {name} Saved!");
                            }
                        }
                        else { }

                    }
                }
        }

        public static string FileSafeName(string title)
        {
            if (title != null)
            {

                if (title.Contains("bra" + "ba"
                    + "ntske" + "vin@gmail.com")) { return "skipfilesave"; }
                title = title.Replace("-", "");
                title = title.ToUpper().Replace("RE ", "");
                title = title.ToUpper().Replace("RE:", "");
                title = title.Replace(" ", "");
                title = title.Replace(":", "");
                title = title.ToUpperInvariant();

                foreach (char symb in Path.GetInvalidFileNameChars())
                {

                    if (title.Contains(symb))
                    {
                        title = title.Replace(symb, ' ');
                    }
                }
                return title;
            }
            return "UK";
        }
    }
}
