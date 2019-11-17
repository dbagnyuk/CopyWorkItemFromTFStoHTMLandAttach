using Microsoft;
using Microsoft.TeamFoundation.Client;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.TeamFoundation.Work;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Controls;
using System;
using System.Net;
using System.IO;
//using System.Windows;
using System.Windows.Forms;
//using System.Windows.Input;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
//using Microsoft.Office.Interop.OneNote;


namespace CopyWorkItemFromTFStoHTMLandAttach
{
    class Program
    {
        public const string Key = "nvision";

        [STAThread]
        static void Main(string[] args)
        {
            Console.Title = "CopyWorkItemFromTFStoHTMLandAttach";

            string configFile = @"CopyWorkItemFromTFStoHTMLandAttach.conf";
            int itemId;
            string[] config = null; // string array for read the config

            while(true)
            {
                if (readConfigFile(configFile, ref config))
                    break;
            }
            while (true)
            {
                Console.WriteLine("Enter TFS number or 'config' for modify [login/password/path to attach] in config file.\n");
                Console.Write("Input TFS id (or 'config'): ");
                //int itemId = 1936268;
                itemId = processingInput();

                if (itemId == 0)
                {
                    editConfigFile(configFile);
                    continue;
                }
                break;
            }

            // save the configuration
            string DomainName = config[0];
            //string Password = config[1];
            string Password = Cipher.Decrypt(config[1], Key);
            //string pathToTasks = config[2];
            string pathToTasks = (config[2].EndsWith("\\") ? config[2] : config[2] + "\\");

            // ask if user want to download the attachments
            Console.Clear();
            Console.Write("Download the Attachments? (y/n): ");
            bool confirm = downloadConfirm();
            Console.Clear();
            Console.Write("Please wait...");

            // create the connection to the TFS server
            NetworkCredential netCred = new NetworkCredential(DomainName, Password);
            Microsoft.VisualStudio.Services.Common.WindowsCredential winCred = new Microsoft.VisualStudio.Services.Common.WindowsCredential(netCred);
            VssCredentials vssCred = new VssCredentials(winCred);
            TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(new Uri("https://tfs.mtsit.com/STS/"), vssCred);

            // catch the authentication error
            try
            {
                tpc.Authenticate();
            }
            catch (Exception ex)
            {
                exExit(ex);
            }

            WorkItemStore workItemStore = tpc.GetService<WorkItemStore>();
            WorkItem workItem = null;

            // catch not existed TFS id
            try
            {
                workItem = workItemStore.GetWorkItem(itemId);
            }
            catch (Exception ex)
            {
                exExit(ex);
            }

            // create web link for tfs id
            string tfsLink = tpc.Uri + workItem.AreaPath.Remove(workItem.AreaPath.IndexOf((char)92)) + "/_workitems/edit/";

            string pathToHtml = pathToTasks + workItem.Type.Name + " " + workItem.Id + ".html";
            string pathToAttach = pathToTasks + workItem.Id;

            FileStream fileStream = null;
            StreamWriter streamWriter = null;

            // create/open the html file
            if (File.Exists(pathToHtml))
                fileStream = new FileStream(pathToHtml, FileMode.Truncate);
            else
                fileStream = new FileStream(pathToHtml, FileMode.CreateNew);
            streamWriter = new StreamWriter(fileStream);

            // fill in the html file
            streamWriter.WriteLine("{0}", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01//EN\" \"http://www.w3.org/TR/html4/strict.dtd\">");
            streamWriter.WriteLine("{0}", "<html>");
            streamWriter.WriteLine("<head>{0}</head>", "<meta charset=\"UTF-8\">");
            streamWriter.WriteLine("<title>{0} {1}</title>", workItem.Type.Name, workItem.Id);
            streamWriter.WriteLine("{0}", "<body>");
            streamWriter.WriteLine("{0}", "");

            streamWriter.WriteLine(@"<p><font style=""background-color:rgb(255, 255, 255); color:rgb(0, 0, 0); font-family:Segoe UI; font-size:12px;"">"
                                   + workItem.Type.Name + " " + workItem.Id + ": " + workItem.Title
                                   + @"</font><p>");

            streamWriter.WriteLine(@"<p style=""border: 1px solid; color: red; width: 50%;"">"
                                   + @"<font style=""background-color:rgb(255, 255, 255); color:rgb(0, 0, 0); font-family:Segoe UI; font-size:12px;"">"
                                   + workItem.Type.Name + " is <b>" + workItem.State + "</b> and Assigned To <b>" + workItem.Fields["Assigned To"].Value + "</b>"
                                   + @"</font><p>");

            streamWriter.WriteLine(@"<div style=""border: 1px solid black; background-color:lightgray;"">TITLE:</div>");
            streamWriter.WriteLine("<p>{0}</p>", workItem.Title);

            streamWriter.WriteLine(@"<div style=""border: 1px solid black; background-color:lightgray;"">DESCRIPTION:</div>");
            if (workItem.Type.Name == "Bug" || workItem.Type.Name == "Issue")
                streamWriter.WriteLine(workItem.Fields["REPRO STEPS"].Value);
            else if (workItem.Type.Name == "Task")
                streamWriter.WriteLine(workItem.Fields["DESCRIPTION"].Value);

            streamWriter.WriteLine(@"<div style=""border: 1px solid black; background-color:lightgray;"">HISTORY:</div><br>");
            for (int i = workItem.Revisions.Count - 1; i >= 0; i--)
            {
                streamWriter.WriteLine(@"<font style=""background-color:rgb(255, 255, 255); color:rgb(0, 0, 0); font-family:Segoe UI; font-size:12px; font-weight:bold;"">"
                                       + workItem.Revisions[i].Fields["Changed By"].Value
                                       + @"</font><br>");
                if (workItem.Revisions[i].Fields["History"].Value.Equals(""))
                    streamWriter.WriteLine(workItem.Revisions[i].Fields["History"].Value);
                else
                    streamWriter.WriteLine(workItem.Revisions[i].Fields["History"].Value
                                           + "<br>");
                streamWriter.WriteLine(@"<font style=""background-color:rgb(255, 255, 255); color:rgb(128, 128, 128); font-family:Segoe UI; font-size:12px;"">"
                                       + "&nbsp;"
                                       + workItem.Revisions[i].Fields["State Change Date"].Value
                                       + @"</font><br><br>");
            }

            streamWriter.WriteLine(@"<div style=""border: 1px solid black; background-color:lightgray;"">ALL LINKS:</div>");
            streamWriter.WriteLine(@"<p><table style=""width:100%; font-family:Segoe UI; font-size:12px;"">");
            streamWriter.WriteLine(@"<tr><th align=""left"">Link Type</th>
                                         <th align=""left"">Work Item Type</th>
                                         <th align=""left"">ID</th>
                                         <th align=""left"">State</th>
                                         <th align=""left"">Title</th>
                                         <th align=""center"">Assigned To</th></tr>");
            foreach (WorkItemLink link in workItem.WorkItemLinks)
            {
                WorkItem wiDeliverable = workItemStore.GetWorkItem(link.TargetId);
                streamWriter.WriteLine(@"<tr><td>{0}</td>", link.LinkTypeEnd.Name);
                streamWriter.WriteLine(@"<td>{0}</td>", wiDeliverable.Type.Name);
                streamWriter.WriteLine(@"<td><a href=""{0}{1}"">{1}</a></td>", tfsLink, wiDeliverable.Id);
                streamWriter.WriteLine(@"<td>{0}</td>", wiDeliverable.State);
                streamWriter.WriteLine(@"<td>{0}</td>", wiDeliverable.Title);
                streamWriter.WriteLine(@"<td>{0}</td></tr>", wiDeliverable.Fields["Assigned To"].Value);
            }
            streamWriter.WriteLine(@"</table></p>");

            streamWriter.WriteLine(@"<div style=""border: 1px solid black; background-color:lightgray;"">LINK:</div>");
            streamWriter.WriteLine(@"<p><a href=""{0}{1}"">{0}{1}</a><p>", tfsLink, workItem.Id);

            // download the attachments from tfs item
            if (confirm)
            {
                // create the path to directory for saving attachments and search if the dir alredy exist
                DirectoryInfo hdDirectoryInWhichToSearch = new DirectoryInfo(pathToTasks);
                FileSystemInfo[] filesAndDirs = hdDirectoryInWhichToSearch.GetFileSystemInfos("*" + workItem.Id + "*");

                foreach (FileSystemInfo foundDir in filesAndDirs)
                    if (foundDir.GetType() == typeof(DirectoryInfo))
                        pathToAttach = foundDir.FullName;

                // if not exists, create it
                if (!Directory.Exists(pathToAttach))
                    Directory.CreateDirectory(pathToAttach);

                // Get a WebClient object to do the attachment download
                WebClient webClient = new WebClient()
                {
                    UseDefaultCredentials = true
                };

                // Loop through each attachment in the work item.
                foreach (Attachment attachment in workItem.Attachments)
                {
                    // Construct a filename for the attachment
                    string filename = string.Format("{0}\\{1}", pathToAttach, attachment.Name);
                    // Download the attachment.
                    webClient.DownloadFile(attachment.Uri, filename);
                    Console.Write(".");
                }

                streamWriter.WriteLine(@"<div style=""border: 1px solid black; background-color:lightgray;"">ATTACHMENTS:</div>");
                streamWriter.WriteLine(@"<p><a href=""{0}"">{0}</a><p>", pathToAttach);
            }

            streamWriter.WriteLine("{0}", "</body>");
            streamWriter.WriteLine("{0}", "</html>");

            streamWriter.Close();
            fileStream.Close();

            // open the created html file, will be open by default app for html files
            System.Diagnostics.Process.Start(pathToHtml);
        }

        // get the input tfs id or 'config' word for go to edit the config
        public static int processingInput()
        {
            int result; // for return
            char enteredSymbol; // char for the value of the pressed key
            int stringSize = 7; // size for char array in which we will store the entered digits
            var sString = new char[stringSize]; // char array in which we will store the entered digits
            int digitsCount = 0; // how many digits we already store

            // loop in which we will analyze pressed keys
            while (true)
            {
                // read the ASCII code from pressed button and save it in char. ReadKey(true) - is for decline write it in console
                enteredSymbol = Console.ReadKey(true).KeyChar;

                // if pressed Ctrl+V we pasted from Clipboard
                if (enteredSymbol == 22 && digitsCount == 0)
                {
                    // cheking if it's the number if else do nothing
                    if (Int32.TryParse(Clipboard.GetText(), out result))
                    {
                        Console.Write(result);
                        Thread.Sleep(500);
                        break;
                    }
                    else
                        continue;
                }

                // condition: if the entered symbol is a later between (a and z) or (A and Z)
                // and the count of entered digits less than the size of the char array
                if ((enteredSymbol >= 48 && enteredSymbol <= 57) && digitsCount < stringSize)
                {
                    Console.Write((char)enteredSymbol);
                    sString[digitsCount++] = (char)enteredSymbol;
                }

                // condition: if the entered symbol is a digit between 0 and 9
                // and the count of entered digits less than the size of the char array
                if (((enteredSymbol >= 65 && enteredSymbol <= 90) || (enteredSymbol >= 97 && enteredSymbol <= 122)) && digitsCount < stringSize - 1)
                {
                    Console.Write((char)enteredSymbol);
                    sString[digitsCount++] = (char)enteredSymbol;
                }
                // condition: if pressed Enter and we nothing entered before, we do nothing
                else if (enteredSymbol == 13 && digitsCount == 0)
                {
                    continue;
                }
                // condition: if pressed Backspace we delete digit from char array which is we wrote in before
                else if (enteredSymbol == 8 && digitsCount > 0)
                {
                    digitsCount--;
                    Console.Write("\b \b"); // return cursor on the previous position
                    sString[digitsCount] = '\0';
                }
                // condition: if pressed Esc we clear the char array and digits in console
                else if (enteredSymbol == 27)
                {
                    // clear the char array from digits we entered before
                    for (int i = 0; i < digitsCount; i++)
                    {
                        sString[i] = '\0';
                        Console.Write("\b \b"); // return cursor on the previous position
                    }
                    digitsCount = 0; // count of digits we entered should be 0
                }
                // condition: if the Enter button pressed we finish read the entering
                else if (enteredSymbol == 13)
                {
                    string savedNumber = new string(sString);
                    if (Int32.TryParse(savedNumber, out result))
                        break;
                    if (savedNumber.ToLower().CompareTo("config") == 0)
                    {
                        result = 0;
                        break;
                    }
                    if (!Int32.TryParse(savedNumber, out result) || savedNumber.ToLower().CompareTo("config") != 0)
                    {
                        // clear the char array from digits we entered before
                        for (int i = 0; i < digitsCount; i++)
                        {
                            sString[i] = '\0';
                            Console.Write("\b \b"); // return cursor on the previous position
                        }
                        digitsCount = 0; // count of digits we entered should be 0
                        continue;
                    }
                }
            }

            Console.Clear(); // clearing the console
            return result;
        }
        // process password input
        public static string passwordInput()
        {
            char enteredSymbol; // char for the value of the pressed key
            int stringSize = 64; // size for char array in which we will store the entered digits
            var sString = new char[stringSize]; // char array in which we will store the entered digits
            int digitsCount = 0; // how many digits we already store
            string password = null; // returned string

            while (true)
            {
                enteredSymbol = Console.ReadKey(true).KeyChar;

                // if pressed Ctrl+V we pasted from Clipboard
                if (enteredSymbol == 22 && digitsCount == 0)
                {
                    password = Clipboard.GetText();
                    Console.Write(password);
                    Thread.Sleep(400);
                    for (int i = 0; i < password.Length; i++)
                        Console.Write("\b \b");
                    for (int i = 0; i < password.Length; i++)
                        Console.Write((char)42);
                    break;
                }
                // control inputted symbol letters, digits, special
                if ((enteredSymbol >= 33 && enteredSymbol <= 126) && digitsCount < stringSize)
                {
                    Console.Write((char)enteredSymbol);
                    Thread.Sleep(400);
                    Console.Write("\b \b");
                    Console.Write((char)42);
                    sString[digitsCount++] = (char)enteredSymbol;
                }

                // condition: if pressed Enter and we nothing entered before, we do nothing
                if (enteredSymbol == 13 && digitsCount == 0)
                {
                    continue;
                }
                // condition: if pressed Backspace we delete digit from char array which is we wrote in before
                else if (enteredSymbol == 8 && digitsCount > 0)
                {
                    digitsCount--;
                    Console.Write("\b \b"); // return cursor on the previous position
                    sString[digitsCount] = '\0';
                }
                // condition: if pressed Esc we clear the char array and digits in console
                else if (enteredSymbol == 27)
                {
                    // clear the char array from digits we entered before
                    for (int i = 0; i < digitsCount; i++)
                    {
                        sString[i] = '\0';
                        Console.Write("\b \b"); // return cursor on the previous position
                    }
                    digitsCount = 0; // count of digits we entered should be 0
                }
                // condition: if the Enter button pressed we finish read the entering
                else if (enteredSymbol == 13)
                {
                    password = new string(sString, 0, digitsCount);
                    break;
                }
            }
            return password;
        }
        // process login and path input
        public static string pathAndLoginInput()
        {
            char enteredSymbol; // char for the value of the pressed key
            int stringSize = 64; // size for char array in which we will store the entered digits
            var sString = new char[stringSize]; // char array in which we will store the entered digits
            int digitsCount = 0; // how many digits we already store
            string output = null;

            while (true)
            {
                enteredSymbol = Console.ReadKey(true).KeyChar;

                // if pressed Ctrl+V we pasted from Clipboard
                if (enteredSymbol == 22 && digitsCount == 0)
                {
                    output = Clipboard.GetText();
                    Console.Write(output);
                    break;
                }
                // control inputted symbols. only letters, digits, colon, and slash
                if (((enteredSymbol >= 48 && enteredSymbol <= 58) || (enteredSymbol >= 65 && enteredSymbol <= 90) || (enteredSymbol >= 97 && enteredSymbol <= 122) || enteredSymbol == 92) && digitsCount < stringSize - 1)
                {
                    Console.Write((char)enteredSymbol);
                    sString[digitsCount++] = (char)enteredSymbol;
                }
                // condition: if pressed Enter and we nothing entered before, we do nothing
                if (enteredSymbol == 13 && digitsCount == 0)
                {
                    sString[digitsCount++] = (char)35;
                    output = new string(sString);
                    break;
                }
                // condition: if pressed Backspace we delete digit from char array which is we wrote in before
                else if (enteredSymbol == 8 && digitsCount > 0)
                {
                    digitsCount--;
                    Console.Write("\b \b"); // return cursor on the previous position
                    sString[digitsCount] = '\0';
                }
                // condition: if pressed Esc we clear the char array and digits in console
                else if (enteredSymbol == 27)
                {
                    // clear the char array from digits we entered before
                    for (int i = 0; i < digitsCount; i++)
                    {
                        sString[i] = '\0';
                        Console.Write("\b \b"); // return cursor on the previous position
                    }
                    digitsCount = 0; // count of digits we entered should be 0
                }
                // condition: if the Enter button pressed we finish read the entering
                else if (enteredSymbol == 13)
                {
                    output = new string(sString, 0, digitsCount);
                    break;
                }
            }
            return output;
        }
        // process the download attach confirmation
        public static bool downloadConfirm()
        {
            char enteredSymbol; // char for the value of the pressed key
            bool confirm = false;

            // loop in which we will analyze pressed keys
            while (true)
            {
                // read the ASCII code from pressed button and save it in char. ReadKey(true) - is for decline write it in console
                enteredSymbol = Console.ReadKey(true).KeyChar;

                // condition: if the entered symbol 'y' or 'Y'
                if ((enteredSymbol == 89 || enteredSymbol == 121))
                {
                    Console.Write((char)enteredSymbol);
                    Thread.Sleep(500);
                    confirm = true;
                    break;
                }

                // condition: if the entered symbol 'n' or 'N'
                if ((enteredSymbol == 78 || enteredSymbol == 110))
                {
                    Console.Write((char)enteredSymbol);
                    Thread.Sleep(500);
                    confirm = false;
                    break;
                }

                continue;
            }

            Console.Clear(); // clearing the console
            return confirm; // return the bool
        }
        // read the config file to memory
        public static bool readConfigFile(string configFile, ref string[] config)
        {
            bool configReaded = false;

            // check if the config file exist and if it more than 0 bytes
            if (!File.Exists(configFile) || new FileInfo(configFile).Length == 0)
            {
                callEditConfig(configFile);
                return configReaded;
            }
            // read the config file
            try
            {
                // read config file into string array
                config = System.IO.File.ReadAllLines(configFile);
            }
            catch (Exception ex)
            {
                exExit(ex);
                callEditConfig(configFile);
                return configReaded;
            }
            // chek config if it has less or more than 3 string
            if (config.Length != 3)
            {
                callEditConfig(configFile);
                return configReaded;
            }
            // check the string from string array if it's empty or not
            foreach (var s in config)
                if (s.Length == 0)
                {
                    callEditConfig(configFile);
                    return configReaded;
                }
            return true;
        }
        // call the creating or editing the onfig file
        public static void callEditConfig(string conf)
        {
            Console.Clear();
            Console.WriteLine("You need to create/edit the config file first!");
            Thread.Sleep(1000);
            Console.Clear();
            editConfigFile(conf);
            Console.Clear();
        }
        // function for create and editing the config file
        public static void editConfigFile(string ConfFile)
        {
            string[] config = null;
            string tempInput;

            FileStream fileStream = null;
            StreamWriter streamWriter = null;

            // check if config exists or biger than 0 byes
            if (File.Exists(ConfFile) && new FileInfo(ConfFile).Length != 0)
            {
                config = System.IO.File.ReadAllLines(ConfFile);
                fileStream = new FileStream(ConfFile, FileMode.Truncate);
            }
            else
            {
                File.Delete(ConfFile);
                fileStream = new FileStream(ConfFile, FileMode.CreateNew);
            }
            streamWriter = new StreamWriter(fileStream);

            Console.WriteLine("Enter a new login, password, and path to attachments.\nUse Enter to leave previous login and path.\n");

            Console.Write("Enter new login [domain\\login]: ");
            tempInput = pathAndLoginInput();
            if (tempInput.CompareTo("#") == 0)
            {
                streamWriter.WriteLine(config[0]);
                Console.Write("\n");
            }
            else
            {
                streamWriter.WriteLine(tempInput.ToLower());
                Console.Write("\n");
            }

            Console.Write("Enter new password: ");
            //streamWriter.WriteLine(passwordInput());
            streamWriter.WriteLine(Cipher.Encrypt(passwordInput(), Key));
            Console.Write("\n");

            //while (true)
            //{
            //    Console.Clear();
            //    Console.Write("Enter old password: ");
            //    string old_pass = Console.ReadLine();
            //    if (config[1].CompareTo(old_pass) == 0)
            //    {
            //        Console.Clear();
            //        Console.Write("Enter new password: ");
            //        config[1] = Console.ReadLine();
            //        break;
            //    }
            //}

            Console.Write("Enter new save destination: ");
            tempInput = pathAndLoginInput();
            if (tempInput.CompareTo("#") == 0)
            {
                streamWriter.WriteLine(config[2]);
                Console.Write("\n");
            }
            else
            {
                streamWriter.WriteLine(tempInput.ToLower());
                Console.Write("\n");
            }

            streamWriter.Close();
            fileStream.Close();

            Console.Clear();
            Console.Write("Config was updated!");
            Thread.Sleep(1000);
        }
        // for write the catched exception and exit
        public static void exExit(Exception ex)
        {
            Console.Clear();
            Console.WriteLine(ex.Message);
            Console.Write("\nPlease 'Enter' for exit...");
            Console.Read();
            System.Environment.Exit(1);
        }
    }
}