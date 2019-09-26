using CommandLine;
using Google.Apis.Admin.Reports.reports_v1;
using Google.Apis.Admin.Reports.reports_v1.Data;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml;
using DriveFile = Google.Apis.Drive.v3.Data.File;

namespace DriveRestorer
{
    class Options
    {
        [Option("dry", Default = false, Required = false, HelpText = "Do a dry run. No files will actually be moved back.")]
        public bool IsDryRun { get; set; }

        [Option("ip", Default = null, Required = false, HelpText = "Optional IP Address that the original move requests came from.")]
        public string IpAddress { get; set; }

        [Option("drive-id", Required = true, HelpText = "The Drive ID.")]
        public string DriveId { get; set; }

        [Option("move-start", Required = true, HelpText = "The date and time when the movements began. Format yyyy-MM-dd HH-mm-ss")]
        public DateTime MoveStart { get; set; }

        [Option("move-end", Required = true, HelpText = "The date and time when the movements finished. Format yyyy-MM-dd HH-mm-ss")]
        public DateTime MoveEnd { get; set; }
    }

    class Program
    {
        static string ApplicationName = "Drive Restore Fixer";

        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args).WithParsed(options =>
            {
                using (ReportsService reportService = GetReportsService())
                {
                    var activities = GetMoveActivities(reportService, options.DriveId, options.MoveStart, options.MoveEnd, options.IpAddress);
                    var moveInfoList = activities.Select(a => new MoveInfo(a)).ToList();
                    List<IGrouping<string, MoveInfo>> folderGroups = moveInfoList.GroupBy(mi => mi.OriginalFolderId).ToList();

                    using (DriveService driveService = GetDriveService())
                    {
                        var totalMovedFiles = moveInfoList.Count;
                        Console.WriteLine($"Found {totalMovedFiles} files that were moved to the drive root between {options.MoveStart} and {options.MoveEnd}{(options.IpAddress != null ? $" by {options.IpAddress}" : null)}.");
                        Console.WriteLine($"Would you like to move them back to their original folders?{(options.IsDryRun ? " (DRY RUN - no files will actually be moved)" : null)} [y/n] ");
                        if (Console.ReadKey(true).Key != ConsoleKey.Y)
                        {
                            return;
                        }

                        var filesMovedBack = 0;
                        for (int i = 0; i < folderGroups.Count; i++)
                        {
                            var folderGroup = folderGroups[i];
                            Console.WriteLine($"{DateTime.Now.ToLongTimeString()}: Folder '{folderGroup.First().OriginalFolderName}' ({folderGroup.Key}) used to contain {folderGroup.Count()} orphaned files. ({i + 1}/{folderGroups.Count}).");
                            foreach (MoveInfo moveInfo in folderGroup)
                            {
                                filesMovedBack++;
                                string percentageCompleted = (filesMovedBack / (double)totalMovedFiles * 100).ToString("0.0");
                                Console.Write($"{DateTime.Now.ToLongTimeString()}: ... Moving {moveInfo.FileName} ({moveInfo.FileId}) to folderId {moveInfo.OriginalFolderId}");
                                DriveFile file = GetFile(driveService, moveInfo.FileId);
                                if (file.Parents.Count == 1 && file.Parents.Single() == options.DriveId)
                                {
                                    if (!options.IsDryRun)
                                    {
                                        ReverseMove(driveService, file, moveInfo);
                                    }
                                    else
                                    {
                                        Console.Write($"[DRY RUN - NOT MOVED]");
                                    }

                                    Console.WriteLine($" [Completed {percentageCompleted}%]");
                                }
                                else
                                {
                                    Console.WriteLine($" [SKIPPED - FILE NO LONGER IN ROOT! {percentageCompleted}%]");
                                }
                            }
                        }
                    }
                }
            });
        }

        private static UserCredential GetCredential(string credentialsFilename, string outputTokenFilename, params string[] scopes)
        {
            UserCredential credential;

            using (var stream =
                new FileStream(credentialsFilename, FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = outputTokenFilename;
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, scopes, "user", CancellationToken.None, new FileDataStore(credPath, true)).Result;
                Console.WriteLine($"Token from {credentialsFilename} saved to {credPath}");
            }

            return credential;
        }

        private static DriveService GetDriveService()
        {
            UserCredential credential = GetCredential("drive_credentials.json", "token-drive.json", DriveService.Scope.Drive, DriveService.Scope.DriveMetadata);

            return new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        private static ReportsService GetReportsService()
        {
            UserCredential credential = GetCredential("admin_sdk_credentials.json", "token-reports.json", ReportsService.Scope.AdminReportsAuditReadonly);

            return new ReportsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        private static DriveFile GetFile(DriveService service, string fileId)
        {
            var request = service.Files.Get(fileId);
            request.Fields = "*";
            request.SupportsAllDrives = true;
            return request.Execute();
        }

        private static List<Activity> GetMoveActivities(ReportsService service, string destinationFolder, DateTime fromDate, DateTime toDate, string ipAddress = null)
        {
            ActivitiesResource.ListRequest request = service.Activities.List("all", "drive");
            request.MaxResults = 1000;
            request.EventName = "move";
            request.Filters = $"destination_folder_id=={destinationFolder}";
            request.StartTime = XmlConvert.ToString(fromDate, XmlDateTimeSerializationMode.Utc);
            request.EndTime = XmlConvert.ToString(toDate, XmlDateTimeSerializationMode.Utc);
            request.ActorIpAddress = ipAddress;

            var activities = new List<Activity>();
            do
            {
                var response = request.Execute();
                activities.AddRange(response.Items);
                request.PageToken = response.NextPageToken;
            }
            while (request.PageToken != null);

            return activities;
        }

        private static DriveFile ReverseMove(DriveService service, DriveFile file, MoveInfo moveInfo)
        {
            var updateRequest = service.Files.Update(new DriveFile(), moveInfo.FileId);
            updateRequest.AddParents = moveInfo.OriginalFolderId;
            updateRequest.RemoveParents = file.Parents[0];
            updateRequest.SupportsAllDrives = true;
            return updateRequest.Execute();
        }

        class MoveInfo
        {
            public string OriginalFolderId { get; set; }
            public string OriginalFolderName { get; set; }
            public string FileId { get; set; }
            public string FileName { get; set; }

            public MoveInfo(Activity moveActivity)
            {
                Activity.EventsData eventInfo = moveActivity.Events[0];
                OriginalFolderId = eventInfo.Parameters.Single(p => p.Name == "source_folder_id").MultiValue.Single();
                OriginalFolderName = eventInfo.Parameters.Single(p => p.Name == "source_folder_title").MultiValue.Single();
                FileId = eventInfo.Parameters.Single(p => p.Name == "doc_id").Value;
                FileName = eventInfo.Parameters.Single(p => p.Name == "doc_title").Value;
            }
        }
    }
}
