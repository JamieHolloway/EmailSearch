using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace EmailSearch
{
    class Program
    {
        static void Main(string[] args)
        {
            var pattern = @"(?i)this is a hack";
            string[] folders = { "Inbox", "todo", "fromJamie", "wip", "Deleted Items", "archive", "Sent Items", "AzureDevOps", "OpticalDRI", "bucket", "ICM" };

            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.ForegroundColor = ConsoleColor.White;
            Console.BufferWidth = 180;
            Console.WindowHeight = 50;
            Console.WindowWidth = 180;
            Console.Title = "Email Power Search";

            switch (args.Length)
            {
                case 1:
                {
                    if (string.IsNullOrEmpty(pattern))
                    {
                        Console.ForegroundColor = ConsoleColor.White;
                        Console.WriteLine("null console input");
                        Environment.Exit(999);
                    }
                    break;
                }
                case 2:
                    folders = new string[] { args[1] };
                    Console.WriteLine($"searching only folder {args[1]}");
                    break;
                default:
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("args?");
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.Read();
                    Environment.Exit(999);
                    break;
            }

            pattern = args[0];
            Console.WriteLine($"RegEx entered was \"{pattern}\"");

            //var outlookApplication = new Application();
            //var outlookNameSpace = outlookApplication.GetNamespace("MAPI");
            var outputFile = $@"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\SearchResult.html";
            var file = new StreamWriter(outputFile);
            var notLaunched = true;

            try
            {
                foreach (var f in folders)
                {
                    var outlookApplication = new Application();
                    var outlookNameSpace = outlookApplication.GetNamespace("MAPI");
                    var folder = outlookNameSpace.Folders["jamieho@microsoft.com"].Folders[f];
                    Console.WriteLine();
                    Console.WriteLine($"Folder {f}");
                    file.WriteLine($"<html><head></head><body><h1>{f}</h1></body></html>");

                    foreach (dynamic mailItem in folder.Items)
                    {
                        var searchArea = "";
                        Console.Write('.');
                        if (!(mailItem is MailItem)) continue;
                        try
                        {
                            searchArea = searchArea + mailItem?.To + " ";
                            searchArea = searchArea + mailItem?.Sender.Name + " ";
                            searchArea = searchArea + mailItem?.Subject + " ";
                            searchArea = searchArea + mailItem?.Body + " ";
                        }
                        catch (System.Exception e)
                        {
                            Console.WriteLine();
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine(e.Message);
                            Console.ForegroundColor = ConsoleColor.White;
                        }
                        var m = Regex.Match(searchArea, pattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);
                        if (!m.Success) continue;
                        Console.WriteLine();
                        Console.WriteLine($"To: {mailItem.To} Subject: {mailItem.Subject}");
                        var sb = new StringBuilder();
                        sb.Append("<html><head><body>");
                        sb.Append($@"<p><b>From:</b> {mailItem.SenderName}</p>");
                        sb.Append($@"<p><b>To:</b> {mailItem.To}</p>");
                        sb.Append($@"<p><b>Sent:</b> {mailItem.ReceivedTime}</p>");
                        if (!string.IsNullOrEmpty(mailItem.CC)) sb.Append($@"<p><b>CC:</b> {mailItem.CC}</p>");
                        if (!string.IsNullOrEmpty(mailItem.BCC)) sb.Append($@"<p><b>BCC:</b> {mailItem.BCC}</p>");
                        sb.Append($@"<p><b>Subject:</b> {mailItem.Subject}</p>");
                        sb.Append("</html></head></body>");
                        file.WriteLine(sb);
                        file.WriteLine(mailItem.HTMLBody);
                        if (!notLaunched) continue;
                        System.Diagnostics.Process.Start(outputFile);
                        notLaunched = false;
                    }
                }
            }
            catch (System.Exception e)
            {
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(e.ToString());
            }
            finally
            {
                file.Close();
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine();
                Console.WriteLine("Search Complete");
                Console.Read();
            }
        }
    }
}