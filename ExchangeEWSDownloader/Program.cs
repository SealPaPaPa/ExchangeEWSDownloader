using Microsoft.Exchange.WebServices.Data;
using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Autodiscover;

namespace RcvAndSendMail
{
    class Program
    {
        private static void manual()
        {
            Console.WriteLine("");
            Console.WriteLine("{0} [delegate / impersonation] [Url] [Domain] [User] [Password] [Target] [Mail Count] (YYYY/MM/DD) (Search Keyword)", Process.GetCurrentProcess().ProcessName);
            Console.WriteLine("Sample:");
        }

        private static void DownloadMail(FindItemsResults<Item> elements)
        {
            foreach (EmailMessage element in elements)
            {
                try
                {
                    element.Load(PropertySet.FirstClassProperties);
                    var mail = element as EmailMessage;
                    var text = element.Body.Text;

                    string mailTitle = element.Subject.Replace("\\", "").Replace("＼", "").Replace("/", "")
                        .Replace("*", "").Replace("\"", "").Replace("<", "").Replace(">", "")
                        .Replace("?", "").Replace("|", "").Replace(":", "_");
                    Console.WriteLine(mailTitle);
                    using (StreamWriter writetext = File.AppendText("Mails\\" + mailTitle + ".html"))
                    {
                        writetext.WriteLine(text);
                        writetext.Close();
                    }

                    foreach (FileAttachment item in element.Attachments)
                    {
                        item.Load("Attachments\\" + item.Name);
                    }
                }
                catch (NullReferenceException e)
                {
                    Console.WriteLine(e);
                    continue;
                }
                catch (InvalidCastException e)
                {
                    Console.WriteLine(e);
                    continue;
                }
            }
        }

        private static void Main(string[] args)
        {
            string account = "";
            string password = "";
            string domain = "";
            string target = "";
            string url = "";
            string time = "2021/01/01";
            string keyword = "";
            int mailCount = 20;
            int type = 0;

            if (args.Length < 6)
            {
                manual();
                return;
            }
            if ("delegate".Equals(args[0]))
            {
                type = 1;
            }
            else if ("impersonation".Equals(args[0]))
            {
                type = 2;
            }
            else
            {
                manual();
                return;
            }

            time = "";
            keyword = "";
            if (args.Length >= 8)
            {
                time = args[7];
            }
            if (args.Length >= 9)
            {
                keyword = args[8];
            }
            url = args[1];
            domain = args[2];
            account = args[3];
            password = args[4];
            target = args[5];
            mailCount = Int32.Parse(args[6]);

            if (!url.Contains("http"))
            {
                url = "https://" + url + "/EWS/Exchange.asmx";
            }
            else
            {
                url += "/EWS/Exchange.asmx";
            }

            if (!Directory.Exists("Attachments"))
            {
                Directory.CreateDirectory("Attachments");
            }
            if (!Directory.Exists("Mails"))
            {
                Directory.CreateDirectory("Mails");
            }

            ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;//<= 加入
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;//<= 加入

            ExchangeService es = new ExchangeService(ExchangeVersion.Exchange2010);//版本預設值最新版
            es.Credentials = new WebCredentials(account, password, domain);

            es.Url = new Uri(url); // Server路徑
            List<SearchFilter> searchFilterCollection = new List<SearchFilter>();
            int flag = 0;
            SearchFilter search;
            if (!time.Equals(""))
            {
                search = new SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived,
                    DateTime.ParseExact(time, "yyyy/MM/dd", null));
                searchFilterCollection.Add(search);
                flag = 1;
            }
            if (!keyword.Equals(""))
            {
                search = new SearchFilter.ContainsSubstring(ItemSchema.Body, keyword);
                searchFilterCollection.Add(search);
                flag = 1;
            }
            if (flag == 1)
            {
                search = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                    searchFilterCollection.ToArray());
            }
            else
            {
                search = null;
            }

            // Get Inbox
            int offset = 0;
            ItemView view = new ItemView(mailCount, offset, OffsetBasePoint.Beginning);
            view.PropertySet = PropertySet.FirstClassProperties;
            FindItemsResults<Item> elements = null;
            if (type == 1)
                elements = es.FindItems(new FolderId(WellKnownFolderName.Inbox, target), search, view);
            else if (type == 2)
            {
                es.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, target);
                elements = es.FindItems(WellKnownFolderName.Inbox, search, view);
            }
            else
                return;

            DownloadMail(elements);

            // Get Sendbox
            offset = 0;
            ItemView view1 = new ItemView(mailCount, offset, OffsetBasePoint.Beginning);
            view1.PropertySet = PropertySet.FirstClassProperties;
            if (type == 1)
                elements = es.FindItems(new FolderId(WellKnownFolderName.SentItems, target), search, view);
            else if (type == 2)
            {
                es.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, target);
                elements = es.FindItems(WellKnownFolderName.SentItems, search, view);
            }
            else
                return;

            DownloadMail(elements);
        }
    }
}