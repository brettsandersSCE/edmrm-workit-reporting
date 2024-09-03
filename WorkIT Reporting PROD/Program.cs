using System;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Security;
using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Text;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.IO; // For handling file operations
using CsvHelper; // For handling CSV file creation (you may need to add this package)
using SPFile = Microsoft.SharePoint.Client.File;
using IOFile = System.IO.File;

class Program
{
    static void Main(string[] args)
    {
        bool edmrmintegration = bool.Parse(ConfigurationManager.AppSettings["edmrmintegration"] ?? "false");
        string ticket = string.Empty; // Initialize ticket variable

        if (edmrmintegration)
        {
            // Run the async method synchronously
            ticket = GetTicketAsync().GetAwaiter().GetResult();
            Console.WriteLine($"Retrieved Ticket: {ticket}");
        }
        else
        {
            Console.WriteLine("Skipping authentication since eDMRM integration is disabled.");
        }

        // Configuration settings
        string siteUrl = ConfigurationManager.AppSettings["SPSite"];
        string username = ConfigurationManager.AppSettings["SPUserName"];
        string password = ConfigurationManager.AppSettings["SPPassword"];
        string emailLogin = ConfigurationManager.AppSettings["Email-Login"];
        string emailHost = ConfigurationManager.AppSettings["Email-Host"];
        string emailPort = ConfigurationManager.AppSettings["Email-Port"];
        string emailPassword = ConfigurationManager.AppSettings["Email-Password"];
        bool emailEnableSsl = bool.Parse(ConfigurationManager.AppSettings["Email-EnableSsl"]);
        string emailRecipients = ConfigurationManager.AppSettings["Email-Recipients"];
        int successHoursThreshold = int.Parse(ConfigurationManager.AppSettings["SuccessHoursThreshold"]);
        bool isDebug = bool.Parse(ConfigurationManager.AppSettings["debug"] ?? "false");
        string debugLibrary = ConfigurationManager.AppSettings["debuglibrary"];
        edmrmintegration = bool.Parse(ConfigurationManager.AppSettings["edmrmintegration"] ?? "true");

        var securePassword = new SecureString();
        foreach (char c in password.ToCharArray())
        {
            securePassword.AppendChar(c);
        }

        var credentials = new SharePointOnlineCredentials(username, securePassword);

        using (var context = new ClientContext(siteUrl))
        {
            context.Credentials = credentials;

            // Retrieve the regional settings to get the time zone
            Web web = context.Web;
            context.Load(web);
            context.Load(web.RegionalSettings);
            context.Load(web.RegionalSettings.TimeZone);
            context.ExecuteQuery();

            // Map SharePoint time zone description to .NET time zone ID
            string timeZoneId = MapSharePointTimeZoneToDotNet(web.RegionalSettings.TimeZone.Description);
            TimeZoneInfo sharePointTimeZone = TimeZoneInfo.FindSystemTimeZoneById(timeZoneId);

            // Get the current server time
            DateTime serverTimeNow = TimeZoneInfo.ConvertTime(DateTime.UtcNow, sharePointTimeZone);

            // Load all lists in the site, but limit to debugLibrary if in debug mode
            ListCollection lists = web.Lists;
            if (isDebug)
            {
                context.Load(lists, lsts => lsts.Include(lst => lst.Title, lst => lst.BaseType, lst => lst.BaseTemplate)
                                                 .Where(lst => lst.Title == debugLibrary));
            }
            else
            {
                context.Load(lists, lsts => lsts.Include(lst => lst.Title, lst => lst.BaseType, lst => lst.BaseTemplate));
            }
            context.ExecuteQuery();

            Console.WriteLine("Libraries found in the site:");
            foreach (List list in lists)
            {
                Console.WriteLine($"- {list.Title}");
            }

            int totalSuccessfulFolders = 0;
            int totalAssembledFolders = 0;
            int totalSentToFaoFolders = 0;
            var delayedTransfers = new List<DelayedTransfer>();
            var successfulTransfersList = new List<SuccessfulTransfer>();

            foreach (List list in lists)
            {
                if (list.BaseTemplate == (int)ListTemplateType.DocumentLibrary && list.BaseType == BaseType.DocumentLibrary)
                {
                    if (isDebug && !list.Title.Equals(debugLibrary, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"Skipping library {list.Title} because it's not the debug library.");
                        continue;
                    }

                    try
                    {
                        Console.WriteLine($"Processing library: {list.Title}");

                        context.Load(list, l => l.RootFolder);
                        context.ExecuteQuery();
                        Console.WriteLine("Loading root folder...");
                        context.Load(list.RootFolder, rf => rf.ServerRelativeUrl);
                        context.ExecuteQuery();
                        Console.WriteLine("Root folder loaded.");

                        Console.WriteLine($"Loading server relative URL of root folder...");
                        string serverRelativeUrl = list.RootFolder.ServerRelativeUrl;
                        Console.WriteLine($"Server relative URL: {serverRelativeUrl}");

                        var allItems = new List<ListItem>();
                        CamlQuery camlQuery = new CamlQuery
                        {
                            ViewXml = "<View Scope='RecursiveAll'><RowLimit>5000</RowLimit></View>",
                            FolderServerRelativeUrl = serverRelativeUrl
                        };

                        ListItemCollectionPosition itemPosition = null;
                        int batchNumber = 1;

                        do
                        {
                            camlQuery.ListItemCollectionPosition = itemPosition;
                            ListItemCollection listItemCollection = list.GetItems(camlQuery);
                            context.Load(listItemCollection);
                            context.ExecuteQuery();

                            allItems.AddRange(listItemCollection);
                            itemPosition = listItemCollection.ListItemCollectionPosition;

                            Console.WriteLine($"Fetched {listItemCollection.Count} items in batch {batchNumber}.");

                            if (itemPosition != null)
                            {
                                Console.WriteLine("More items to fetch...");
                                batchNumber++;
                            }
                            else
                            {
                                Console.WriteLine("No more items to fetch.");
                            }

                        } while (itemPosition != null);

                        Console.WriteLine($"Total items fetched: {allItems.Count}");

                        var filteredItems = new List<ListItem>();

                        foreach (var item in allItems)
                        {
                            try
                            {
                                // Try to access the Status and Modified fields directly
                                string status = item["Status"]?.ToString();
                                DateTime modifiedDate = (DateTime)item["Modified"];

                                // Apply your filtering logic here
                                if (status == "ASSEMBLED" ||
                                    status == "Sent to FAO" ||
                                    status == "FCMP WO Status Set" ||
                                    status == "CLSD WO Status Set" ||
                                    (status == "Extraction Success" && (serverTimeNow - TimeZoneInfo.ConvertTimeFromUtc(modifiedDate, TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time"))).TotalHours <= successHoursThreshold))
                                {
                                    filteredItems.Add(item);
                                }
                            }
                            catch (PropertyOrFieldNotInitializedException ex)
                            {
                                Console.WriteLine($"Error: Property or field not initialized for item with ID {item.Id}. Exception: {ex.Message}");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Unexpected error encountered for item with ID {item.Id}. Exception: {ex.Message}");
                            }
                        }

                        Console.WriteLine($"Filtered items count: {filteredItems.Count}");


                        foreach (var item in filteredItems)
                        {
                            string status = item["Status"]?.ToString();
                            DateTime modifiedDate = TimeZoneInfo.ConvertTimeFromUtc((DateTime)item["Modified"], TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time"));
                            double delayHours = (serverTimeNow - modifiedDate).TotalHours;

                            Console.WriteLine($"Item ID: {item.Id}, Status: '{status}', Modified: {modifiedDate}, Delay Hours: {delayHours}");

                            if (status.Equals("Extraction Success", StringComparison.OrdinalIgnoreCase) && delayHours <= successHoursThreshold)
                            {
                                totalSuccessfulFolders++;

                                JObject apiResult = null;
                                if (!isDebug && edmrmintegration)
                                {
                                    apiResult = GetWOInfoAsync(ticket, item["FileLeafRef"].ToString()).GetAwaiter().GetResult();
                                }

                                string wsid = isDebug || apiResult == null ? "Debug Mode" : (apiResult?["data"]?["dataid"]?.ToString() ?? "No eDMRM workspace");
                                int finalPlanningPkg = isDebug ? 0 : apiResult?["data"]?["fpp"]?.ToObject<int>() ?? 0;
                                int finalWOPkg = isDebug ? 0 : apiResult?["data"]?["fwop"]?.ToObject<int>() ?? 0;

                                Folder folder = web.GetFolderByServerRelativeUrl(list.RootFolder.ServerRelativeUrl + "/" + item["FileLeafRef"]);
                                int spFileCount = GetTotalFileCount(folder, context);

                                successfulTransfersList.Add(new SuccessfulTransfer
                                {
                                    LibraryName = list.Title,
                                    FolderName = item["FileLeafRef"].ToString(),
                                    WorkspaceID = wsid,
                                    FinalPlanningPkg = finalPlanningPkg,
                                    FinalWOPkg = finalWOPkg,
                                    SPFileCount = spFileCount,
                                    ExtractionDate = modifiedDate
                                });
                            }
                            else
                            {
                                string wsInfo = "Debug Mode";
                                if (!isDebug && edmrmintegration)
                                {
                                    JObject apiResult = GetWOInfoAsync(ticket, item["FileLeafRef"].ToString()).GetAwaiter().GetResult();
                                    if (apiResult != null && apiResult.ContainsKey("data"))
                                    {
                                        try
                                        {
                                            string dataField = apiResult["data"]?.ToString();
                                            if (!string.IsNullOrEmpty(dataField))
                                            {
                                                JObject parsedData = JObject.Parse(dataField);
                                                int workspaceID = parsedData["dataid"]?.ToObject<int>() ?? 0;
                                                if (workspaceID != 0)
                                                {
                                                    wsInfo = workspaceID.ToString();
                                                }
                                            }
                                        }
                                        catch (JsonReaderException)
                                        {
                                            wsInfo = "No eDMRM workspace";
                                        }
                                    }
                                }

                                delayedTransfers.Add(new DelayedTransfer
                                {
                                    LibraryName = list.Title,
                                    FolderName = item["FileLeafRef"].ToString(),
                                    LastModified = modifiedDate,
                                    PendingHours = delayHours,
                                    SharepointCategory = status,
                                    EDMRMCategory = isDebug ? "Debug Mode" : "Final Planning Package",
                                    FileCount = GetTotalFileCount(web.GetFolderByServerRelativeUrl(list.RootFolder.ServerRelativeUrl + "/" + item["FileLeafRef"]), context),
                                    WorkspaceIDInfo = wsInfo
                                });
                            }
                        }
                    }
                    catch (ServerException ex)
                    {
                        Console.WriteLine($"Error processing library {list.Title}: {ex.Message}");
                    }
                }
            }


            // Sort the delayed transfers by PendingHours in descending order
            delayedTransfers = delayedTransfers.OrderByDescending(dt => dt.PendingHours).ToList();

            var emailBody = new StringBuilder();
            emailBody.AppendLine("<html><body>");
            emailBody.AppendLine("<p style='font-size:16pt;'>WorkIT Sharepoint/eDMRM integration Report</p>");

            // Add the report run time above the first table in small grey font
            string reportRunTime = DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss");
            emailBody.AppendLine($"<p style='font-size:8pt;color:grey;'>Report run time {reportRunTime}</p>");

            // Calculate counts based on the delayedTransfers list
            totalAssembledFolders = delayedTransfers.Count(x => x.SharepointCategory == "ASSEMBLED");
            totalSentToFaoFolders = delayedTransfers.Count(x => x.SharepointCategory == "Sent to FAO");
            int totalFCMPFolders = delayedTransfers.Count(x => x.SharepointCategory == "FCMP WO Status Set");
            int totalCLSDFolders = delayedTransfers.Count(x => x.SharepointCategory == "CLSD WO Status Set");

            // Now update the summary table with these counts

            emailBody.AppendLine("<table style='border:1px solid black;border-collapse:collapse;'>");

            // Row 1: The green status row (spans two columns)
            emailBody.AppendLine("<tr>");
            emailBody.AppendLine("<td style='background-color:green;color:white;padding:10px;text-align:center;' colspan='2'>");
            emailBody.AppendLine($"&nbsp;Successful eDMRM archivals (last {successHoursThreshold} hours): {totalSuccessfulFolders}&nbsp;");
            emailBody.AppendLine("</td>");
            emailBody.AppendLine("</tr>");

            // Row 2: Total pending transfers with borders
            int totalPendingTransfers = totalAssembledFolders + totalSentToFaoFolders + totalFCMPFolders + totalCLSDFolders;
            emailBody.AppendLine("<tr>");
            emailBody.AppendLine("<td style='border:1px solid black;font-weight:bold;'>&nbsp;Total Pending Transfers:&nbsp;</td>");
            emailBody.AppendLine($"<td style='border:1px solid black;'>&nbsp;{totalPendingTransfers}&nbsp;</td>");
            emailBody.AppendLine("</tr>");

            // Subcategory rows for each status type with borders
            emailBody.AppendLine("<tr>");
            emailBody.AppendLine("<td style='border:1px solid black;padding-left:20px;'>&nbsp;Assembled:&nbsp;</td>");
            emailBody.AppendLine($"<td style='border:1px solid black;'>&nbsp;{totalAssembledFolders}&nbsp;</td>");
            emailBody.AppendLine("</tr>");

            emailBody.AppendLine("<tr>");
            emailBody.AppendLine("<td style='border:1px solid black;padding-left:20px;'>&nbsp;Sent to FAO:&nbsp;</td>");
            emailBody.AppendLine($"<td style='border:1px solid black;'>&nbsp;{totalSentToFaoFolders}&nbsp;</td>");
            emailBody.AppendLine("</tr>");

            emailBody.AppendLine("<tr>");
            emailBody.AppendLine("<td style='border:1px solid black;padding-left:20px;'>&nbsp;FCMP WO Status Set:&nbsp;</td>");
            emailBody.AppendLine($"<td style='border:1px solid black;'>&nbsp;{totalFCMPFolders}&nbsp;</td>");
            emailBody.AppendLine("</tr>");

            emailBody.AppendLine("<tr>");
            emailBody.AppendLine("<td style='border:1px solid black;padding-left:20px;'>&nbsp;CLSD WO Status Set:&nbsp;</td>");
            emailBody.AppendLine($"<td style='border:1px solid black;'>&nbsp;{totalCLSDFolders}&nbsp;</td>");
            emailBody.AppendLine("</tr>");

            // Calculate the average pending transfer using the separate function
            double averagePendingTransfer = CalculateAveragePendingTransfer(delayedTransfers);

            // Determine the row style based on the averagePendingTransfer value
            string rowStyle;
            if (averagePendingTransfer <= 24)
            {
                rowStyle = "background-color:green;color:white;font-weight:normal;";
            }
            else if (averagePendingTransfer > 24 && averagePendingTransfer < 48)
            {
                rowStyle = "background-color:yellow;color:black;font-weight:bold;";
            }
            else
            {
                rowStyle = "background-color:red;color:white;font-weight:normal;";
            }

            // Add the Average Pending Transfer row with the determined style
            emailBody.AppendLine("<tr>");
            emailBody.AppendLine($"<td style='border:1px solid black;font-weight:bold;{rowStyle}'>&nbsp;Average Pending Transfer:&nbsp;</td>");
            emailBody.AppendLine($"<td style='border:1px solid black;{rowStyle}'>&nbsp;{averagePendingTransfer:F2} hours&nbsp;</td>");
            emailBody.AppendLine("</tr>");




            emailBody.AppendLine("</table>");



            // Add a break before the next section
            emailBody.AppendLine("<br>");

            // Next section: Pending Transfers (listing each delayed transfer)
            emailBody.AppendLine("<p style='font-weight:bold;'>Pending Transfers:</p>");
            emailBody.AppendLine("<table style='border:1px solid black;border-collapse:collapse;'>");
            emailBody.AppendLine("<tr style='background-color:red;color:white;'>");
            emailBody.AppendLine("<th style='border:1px solid black;'>&nbsp;Library Name&nbsp;</th>");
            emailBody.AppendLine("<th style='border:1px solid black;'>&nbsp;Folder Name&nbsp;</th>");
            emailBody.AppendLine("<th style='border:1px solid black;'>&nbsp;Last Modified (PST)&nbsp;</th>");
            emailBody.AppendLine("<th style='border:1px solid black;'>&nbsp;Pending (hours)&nbsp;</th>");
            emailBody.AppendLine("<th style='border:1px solid black;'>&nbsp;Sharepoint Category&nbsp;</th>");
            emailBody.AppendLine("<th style='border:1px solid black;'>&nbsp;eDMRM Category&nbsp;</th>");
            emailBody.AppendLine("<th style='border:1px solid black;'>&nbsp;FileCount (SP)&nbsp;</th>");
            emailBody.AppendLine("<th style='border:1px solid black;'>&nbsp;eDMRM WSID&nbsp;</th>");
            emailBody.AppendLine("</tr>");

            foreach (var transfer in delayedTransfers)
            {
                emailBody.AppendLine($"<tr><td style='border:1px solid black;'>&nbsp;{transfer.LibraryName}&nbsp;</td><td style='border:1px solid black;'>&nbsp;{transfer.FolderName}&nbsp;</td><td style='border:1px solid black;'>&nbsp;{transfer.LastModified:yyyy-MM-dd HH:mm}&nbsp;</td><td style='border:1px solid black;'>&nbsp;{transfer.PendingHours:F2}&nbsp;</td><td style='border:1px solid black;'>&nbsp;{transfer.SharepointCategory}&nbsp;</td><td style='border:1px solid black;'>&nbsp;{transfer.EDMRMCategory}&nbsp;</td><td style='border:1px solid black;'>&nbsp;{transfer.FileCount}&nbsp;</td><td style='border:1px solid black;'>&nbsp;{transfer.WorkspaceIDInfo}&nbsp;</td></tr>");
            }

            emailBody.AppendLine("</table>");

            // Add a break before the next section
            emailBody.AppendLine("<br>");

            // Next section: Successful Transfers
            emailBody.AppendLine("<p style='font-weight:bold;'>Successful Transfers:</p>");
            emailBody.AppendLine("<table style='border:1px solid black;border-collapse:collapse;'>");
            emailBody.AppendLine("<tr style='background-color:green;color:white;'>");
            emailBody.AppendLine("<th style='border:1px solid black;'>&nbsp;Library Name&nbsp;</th>");
            emailBody.AppendLine("<th style='border:1px solid black;'>&nbsp;Folder Name&nbsp;</th>");
            emailBody.AppendLine("<th style='border:1px solid black;'>&nbsp;SP FileCount&nbsp;</th>");
            emailBody.AppendLine("<th style='border:1px solid black;'>&nbsp;WorkspaceID&nbsp;</th>");
            emailBody.AppendLine("<th style='border:1px solid black;'>&nbsp;Final Planning Pkg&nbsp;</th>");
            emailBody.AppendLine("<th style='border:1px solid black;'>&nbsp;FinalWOPkg&nbsp;</th>");
            emailBody.AppendLine("<th style='border:1px solid black;'>&nbsp;Extraction Date (PST)&nbsp;</th>"); // New column for extraction date
            emailBody.AppendLine("</tr>");

            foreach (var transfer in successfulTransfersList)
            {
                emailBody.AppendLine($"<tr><td style='border:1px solid black;'>&nbsp;{transfer.LibraryName}&nbsp;</td>");
                emailBody.AppendLine($"<td style='border:1px solid black;'>&nbsp;{transfer.FolderName}&nbsp;</td>");
                emailBody.AppendLine($"<td style='border:1px solid black;'>&nbsp;{transfer.SPFileCount}&nbsp;</td>");
                emailBody.AppendLine($"<td style='border:1px solid black;'>&nbsp;{transfer.WorkspaceID}&nbsp;</td>");
                emailBody.AppendLine($"<td style='border:1px solid black;'>&nbsp;{transfer.FinalPlanningPkg}&nbsp;</td>");
                emailBody.AppendLine($"<td style='border:1px solid black;'>&nbsp;{transfer.FinalWOPkg}&nbsp;</td>");
                emailBody.AppendLine($"<td style='border:1px solid black;'>&nbsp;{transfer.ExtractionDate:yyyy-MM-dd HH:mm}&nbsp;</td></tr>"); // Display extraction date
            }

            emailBody.AppendLine("</table>");

            // Create the CSV file
            // Retrieve the environment variable from the config file
            string environment = ConfigurationManager.AppSettings["environment"] ?? "UnknownEnv";

            // Generate the timestamp and construct the filename
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileName = $"WorkIT_{environment}_{timestamp}.csv";

            // Ensure TMPfiles folder exists
            string tmpFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMPfiles");
            if (!Directory.Exists(tmpFolderPath))
            {
                Directory.CreateDirectory(tmpFolderPath);
            }

            // Delete all CSV files in the TMPfiles folder
            string[] csvFiles = Directory.GetFiles(tmpFolderPath, "*.csv");
            foreach (string file in csvFiles)
            {
                try
                {
                    System.IO.File.Delete(file);
                    Console.WriteLine($"Deleted file: {file}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error deleting file {file}: {ex.Message}");
                }
            }


            // Define the file path within the TMPfiles folder
            string csvFilePath = Path.Combine(tmpFolderPath, fileName);

            // Create the CSV file
            using (var writer = new StreamWriter(csvFilePath))
            using (var csv = new CsvWriter(writer, System.Globalization.CultureInfo.InvariantCulture))
            {
                // Write the headers
                csv.WriteField("Library Name");
                csv.WriteField("Folder Name");
                csv.WriteField("Last Modified (PST)");
                csv.WriteField("Status");
                csv.WriteField("Pending (hours)");
                csv.WriteField("Sharepoint Category");
                csv.WriteField("eDMRM Category");
                csv.WriteField("FileCount (SP)");
                csv.WriteField("eDMRM WSID");
                csv.WriteField("edmrm_FPP_count");
                csv.WriteField("edmrm_FWOP_count");
                csv.WriteField("extraction date");
                csv.NextRecord();

                // Write pending transfers
                foreach (var transfer in delayedTransfers)
                {
                    csv.WriteField(transfer.LibraryName);
                    csv.WriteField(transfer.FolderName);
                    csv.WriteField(transfer.LastModified.ToString("yyyy-MM-dd HH:mm"));
                    csv.WriteField("Pending");
                    csv.WriteField(transfer.PendingHours.ToString("F2"));
                    csv.WriteField(transfer.SharepointCategory);
                    csv.WriteField(transfer.EDMRMCategory);
                    csv.WriteField(transfer.FileCount);
                    csv.WriteField(transfer.WorkspaceIDInfo);
                    csv.WriteField("N/A");
                    csv.WriteField("N/A");
                    csv.WriteField("N/A");
                    csv.NextRecord();
                }

                // Write successful transfers
                foreach (var transfer in successfulTransfersList)
                {
                    csv.WriteField(transfer.LibraryName);
                    csv.WriteField(transfer.FolderName);
                    csv.WriteField(transfer.ExtractionDate.ToString("yyyy-MM-dd HH:mm"));
                    csv.WriteField("Extracted");
                    csv.WriteField("N/A");
                    csv.WriteField("N/A");
                    csv.WriteField("N/A");
                    csv.WriteField(transfer.SPFileCount);
                    csv.WriteField(transfer.WorkspaceID);
                    csv.WriteField(transfer.FinalPlanningPkg);
                    csv.WriteField(transfer.FinalWOPkg);
                    csv.WriteField(transfer.ExtractionDate.ToString("yyyy-MM-dd HH:mm"));
                    csv.NextRecord();
                }
            }

            // Send the email with the CSV attachment
            SendEmail(emailLogin, emailPassword, emailHost, int.Parse(emailPort), emailEnableSsl, emailRecipients, $"WorkIT - eDMRM report for {DateTime.Now:yyyyMMdd_HH:mm:ss}", emailBody.ToString(), csvFilePath);



            // Print the total counts
            Console.WriteLine($"Successful eDMRM archivals (last {successHoursThreshold} hours): {totalSuccessfulFolders}");
            Console.WriteLine($"Total 'Assembled' folders: {totalAssembledFolders}");
            Console.WriteLine($"Total 'Sent to FAO' folders: {totalSentToFaoFolders}");
            Console.WriteLine("\nDelayed Transfers:");
            Console.WriteLine("Library Name, Folder Name, Pending (hours), Sharepoint Category, eDMRM Category, FileCount, eDMRM WSID");
            foreach (var transfer in delayedTransfers)
            {
                Console.WriteLine($"{transfer.LibraryName}, {transfer.FolderName}, {transfer.PendingHours:F2}, {transfer.SharepointCategory}, {transfer.EDMRMCategory}, {transfer.FileCount}, {transfer.WorkspaceIDInfo}");
            }

            Console.WriteLine("\nSuccessful Transfers:");
            Console.WriteLine("Library Name, Folder Name, SP FileCount, WorkspaceID, Final Planning Pkg, FinalWOPkg");
            foreach (var transfer in successfulTransfersList)
            {
                Console.WriteLine($"{transfer.LibraryName}, {transfer.FolderName}, {transfer.SPFileCount}, {transfer.WorkspaceID}, {transfer.FinalPlanningPkg}, {transfer.FinalWOPkg}");
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }

    static int GetTotalFileCount(Folder folder, ClientContext context)
    {
        int maxRetries = int.Parse(ConfigurationManager.AppSettings["edmrmmaxretries"] ?? "10");
        int delaySeconds = int.Parse(ConfigurationManager.AppSettings["retryDelaySeconds"] ?? "60");
        int fileCount = 0;

        for (int attempt = 1; attempt <= maxRetries; attempt++)
        {
            try
            {
                // Load the folder's files and subfolders explicitly
                context.Load(folder, f => f.Files, f => f.Folders);
                context.ExecuteQuery();

                fileCount = folder.Files.Count;

                foreach (Folder subFolder in folder.Folders)
                {
                    // Recursively count files in subfolders
                    fileCount += GetTotalFileCount(subFolder, context);
                }

                return fileCount;
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName.Contains("429"))
                {
                    Console.WriteLine($"429 Too Many Requests encountered. Retrying in {delaySeconds} seconds... (Attempt {attempt}/{maxRetries})");
                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(delaySeconds));
                }
                else
                {
                    Console.WriteLine($"Error encountered: {ex.Message}. Retrying in {delaySeconds} seconds... (Attempt {attempt}/{maxRetries})");
                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(delaySeconds));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error encountered: {ex.Message}. Retrying in {delaySeconds} seconds... (Attempt {attempt}/{maxRetries})");
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(delaySeconds));
            }
        }

        Console.WriteLine("Max retries reached. Operation failed.");
        return fileCount;
    }

    static void SendEmail(string emailLogin, string emailPassword, string emailHost, int emailPort, bool emailEnableSsl, string emailRecipients, string subject, string body, string attachmentFilePath)
    {
        try
        {
            var smtpClient = new SmtpClient(emailHost)
            {
                Port = emailPort,
                Credentials = new NetworkCredential(emailLogin, emailPassword),
                EnableSsl = emailEnableSsl,
            };

            var mailMessage = new MailMessage
            {
                From = new MailAddress(emailLogin),
                Subject = subject,
                Body = body,
                IsBodyHtml = true,
            };

            foreach (var recipient in emailRecipients.Split(';'))
            {
                mailMessage.To.Add(recipient);
            }

            // Attach the CSV file
            if (System.IO.File.Exists(attachmentFilePath))
            {
                mailMessage.Attachments.Add(new System.Net.Mail.Attachment(attachmentFilePath));
            }

            smtpClient.Send(mailMessage);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending email: {ex.Message}");
        }
    }

    static string MapSharePointTimeZoneToDotNet(string sharePointTimeZoneDescription)
    {
        var timeZoneMappings = new Dictionary<string, string>
        {
            { "(UTC-08:00) Pacific Time (US and Canada)", "Pacific Standard Time" },
            { "(UTC-05:00) Eastern Time (US and Canada)", "Eastern Standard Time" },
        };

        if (timeZoneMappings.TryGetValue(sharePointTimeZoneDescription, out string dotNetTimeZoneId))
        {
            return dotNetTimeZoneId;
        }

        throw new TimeZoneNotFoundException($"The time zone ID '{sharePointTimeZoneDescription}' was not found in the mappings.");
    }

    static async Task<string> GetTicketAsync()
    {
        string baseUrl = ConfigurationManager.AppSettings["csurl"];
        string username = ConfigurationManager.AppSettings["username"];
        string password = ConfigurationManager.AppSettings["password"];

        string authUrl = $"{baseUrl.TrimEnd('/')}/api/v1/auth";
        Console.WriteLine($"Starting GetTicketAsync function for URL: {authUrl}");

        string encodedUsername = Uri.EscapeDataString(username);
        string encodedPassword = Uri.EscapeDataString(password);
        string payload = $"username={encodedUsername}&password={encodedPassword}";

        var content = new StringContent(payload);
        content.Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded");

        try
        {
            using (var client = new HttpClient())
            {
                HttpResponseMessage response = await client.PostAsync(authUrl, content);
                response.EnsureSuccessStatusCode();
                string responseBody = await response.Content.ReadAsStringAsync();

                var jsonObj = JObject.Parse(responseBody);
                string otcsticket = jsonObj["ticket"].ToString();

                Console.WriteLine($"Authentication successful. Ticket: {otcsticket}");
                return otcsticket;
            }
        }
        catch (Exception e)
        {
            Console.WriteLine($"Error during authentication, check server connections: {e.Message}");
            throw;
        }
    }


    static double CalculateAveragePendingTransfer(List<DelayedTransfer> delayedTransfers)
    {
        // Filter only the relevant categories for pending transfers
        var relevantPendingTransfers = delayedTransfers.Where(transfer =>
            transfer.SharepointCategory == "ASSEMBLED" ||
            transfer.SharepointCategory == "Sent to FAO" ||
            transfer.SharepointCategory == "FCMP WO Status Set" ||
            transfer.SharepointCategory == "CLSD WO Status Set").ToList();

        // Calculate the total pending hours and count
        double totalPendingHours = relevantPendingTransfers.Sum(transfer => transfer.PendingHours);
        int pendingTransfersCount = relevantPendingTransfers.Count;

        // Calculate and return the average pending transfer
        return pendingTransfersCount > 0 ? totalPendingHours / pendingTransfersCount : 0;
    }


    static async Task<JObject> GetWOInfoAsync(string otcsticket, string folderName)
    {
        // Retrieve the csWR value from the config file
        string csWR = ConfigurationManager.AppSettings["csWR"];
        string csUrl = ConfigurationManager.AppSettings["csurl"];

        // Construct the API endpoint URL
        string url = $"{csUrl}/api/v1/nodes/{csWR}/output?wonumber={folderName}";

        using (var client = new HttpClient())
        {
            // Create the GET request
            var request = new HttpRequestMessage(HttpMethod.Get, url);

            // Add the headers as needed
            request.Headers.Add("otcsticket", otcsticket);

            try
            {
                // Send the request
                HttpResponseMessage response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();

                // Read and parse the JSON response
                string responseBody = await response.Content.ReadAsStringAsync();
                Console.WriteLine(responseBody); // Optional: log the raw response
                return JObject.Parse(responseBody);
            }
            catch (HttpRequestException e)
            {
                Console.WriteLine($"Request error: {e.Message}");
                return null;
            }
        }
    }

    class DelayedTransfer
    {
        public string LibraryName { get; set; }
        public string FolderName { get; set; }
        public DateTime LastModified { get; set; } // Store in PST
        public double PendingHours { get; set; }
        public string SharepointCategory { get; set; }
        public string EDMRMCategory { get; set; }
        public int FileCount { get; set; }
        public string WorkspaceIDInfo { get; set; }
    }

    class SuccessfulTransfer
    {
        public string LibraryName { get; set; }
        public string FolderName { get; set; }
        public string WorkspaceID { get; set; }
        public int FinalPlanningPkg { get; set; }
        public int FinalWOPkg { get; set; }
        public int SPFileCount { get; set; }
        public DateTime ExtractionDate { get; set; } // New property for the modified date
    }
}
