using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailSaving
{
    internal class EmailSorter
    {

        /// <summary>
        /// Seperates all Emails from a collection of items; Only retrieves emails that are newer than the DateTime
        /// </summary>
        /// <param name="Items"></param>
        /// <param name="afterTime"></param>
        /// <returns></returns>
        public static List<MailItem> SeperateEmailsFromItems(Items Items, DateTime beforeTime)
        {
            List<MailItem> mailItems = new List<MailItem>();

            for (int i = Items.Count; i > 0; i--)
            {

                if (Items[i] is MailItem mailItem)
                {
                    if (mailItem.SentOn >= beforeTime)
                        mailItems.Add(mailItem);
                }
            }

            return mailItems;
        }


    }
}
