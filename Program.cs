using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using MailTests;
using System.Net.Mail;
using EmailSaving;
//using Microsoft.Office.Interop.Excel;

namespace MailTests
{
    static class Program
    {
        public static string JSONPATH { get; set; } = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\EmailSaving\\";
        static void Main(string[] args)
        {

            //Helper.CreateCustomerDatabase();


            int PauseTime = 60;
            // Date Time Conversions
            DateTime today = DateTime.Now;
            
            // Used to check if an email came within a certain period
            DateTime check = today.AddDays(-7);
            // Saves all logged in Outlook accounts to search emails from
            Accounts emailAccounts = EmailFilePath.LoggedInOutlookAccounts();

            GeoAccounts accounts = new GeoAccounts();


            Microsoft.Office.Interop.Outlook.Application outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
            NameSpace outlookNamespace;
            MAPIFolder inboxFolder;


            accounts = Helper.AllGeoAccounts();



            foreach (Account acc in emailAccounts)
            {

                outlookApplication = acc.Application;
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);

                EmailFileSaver.SaveAllEmails(accounts, inboxFolder, check,  EmailFileSaver.SaveOutbox);

                outlookApplication = acc.Application;
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                EmailFileSaver.SaveAllEmails(accounts, inboxFolder, check,  EmailFileSaver.SaveInbox);
            }

            outlookApplication = null;


            //emailBox.SaveValues();

            //Console.WriteLine($"Next Save at: {DateTime.Now.AddHours(0.15).ToString()}");

            Thread.Sleep(PauseTime * 60 * 1000);

            Process.Start("S:\\Kevin's Projects\\AutoSearch\\EmailSaving.application");

            Environment.Exit(0);
        }
    }

}



static class Helper
{

    /// <summary>
    /// Saves The xls spread sheet to a json file, the emails need to be on column 4
    /// </summary>

    /*
    public static void CreateCustomerDatabase()
    {
        GeoAccounts Data = new GeoAccounts();

        List<GeoAccount> DatabaseList = new List<GeoAccount>();

        string json;

        string ExcelPath = "C:\\Users\\techsupport117\\Desktop\\EmailSaving\\All mmk.xls";

        string SavePath = "C:\\Users\\techsupport117\\Desktop\\EmailSaving\\Allcustomers.json";

        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
        Workbook wb = excel.Workbooks.Open(ExcelPath);
        Worksheet xlSheet = wb.ActiveSheet;



        for (int i = 2; i < 330000; i++)
        {
            GeoAccount account = new GeoAccount();
            double temp;
            // Console.WriteLine(i);
            if (xlSheet.Cells[i, 1].Value2 == null)
            {

                if (xlSheet.Cells[i, 2].Value2 == null)
                    break;

                continue;

            }



            if (xlSheet.Cells[i, 1].Value2 != null)
            {
                account.Name = Convert.ToString(xlSheet.Cells[i, 1].Value);
            }
            if (xlSheet.Cells[i, 2].Value2 != null)
            {
                temp = xlSheet.Cells[i, 2].Value;
                account.SubId = Convert.ToInt32(temp);
            }
            if (xlSheet.Cells[i, 3].Value2 != null)
            {
                account.AccountNumber = Convert.ToString(xlSheet.Cells[i, 3].Value);
            }
            if (xlSheet.Cells[i, 4].Value2 != null)
            {
                account.Email = Convert.ToString(xlSheet.Cells[i, 4].Value);
                account.Email = account.Email.Replace(" ", "");
            }

            DatabaseList.Add(account);
        }

        Data.list = DatabaseList.ToArray();

        json = JsonSerializer.Serialize<GeoAccounts>(Data);

        File.WriteAllText(SavePath, json);
    }
    */

    public static GeoAccounts AllGeoAccounts()
    {
        string jsonPath = "S:\\Kevin's Projects\\DataHub\\Allcustomers.json";

        string json = File.ReadAllText(jsonPath);

        GeoAccounts accounts = JsonSerializer.Deserialize<GeoAccounts>(json);

        if (accounts != null)
            return accounts;
        else
            return new GeoAccounts();
    }





}


public class GeoAccount
{
    public string Name { get; set; } = "Unkown";
    public int SubId { get; set; } = -1;
    public string AccountNumber { get; set; } = "None";
    public string Email { get; set; } = "Unkown";
}
public class GeoAccounts
{
    public GeoAccount[] list { get; set; }
}

public class EmailCount
{
    public int inbox { get; set; } = 0;
    public int outbox { get; set; } = 0;

    public static EmailCount GetValues()
    {
        EmailCount count = new EmailCount();

        string json = File.ReadAllText(Program.JSONPATH);

        count = JsonSerializer.Deserialize<EmailCount>(json);

        return count;
    }

    public void SaveValues()
    {
        string json = JsonSerializer.Serialize(this);

        Directory.CreateDirectory(Program.JSONPATH);

        File.SetAttributes(Program.JSONPATH, FileAttributes.Normal);

        File.WriteAllText(Program.JSONPATH + "Data.json", json);
    }
}