using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;


namespace CFCopier
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Google Sheets API .NET Quickstart";
        static string SourcePageName = "PageTemplate";
        static string CurrentSpreadSheetName = string.Empty;
        private static int SourcePageId = -1;

        static void Main(string[] args)
        {
            #region Init

            try
            {
                if (!EventLog.SourceExists("CFAppLog"))
                {
                    EventLog.CreateEventSource(new EventSourceCreationData("CFAppLog", "CFAppLog"));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please run from administrator " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var keepLiveMode = false;
            if (args.Length > 0)
            {
                if (args[0].ToUpper() == "KEEPLIVE")
                {
                    LogAnEvent("Keep live mode");
                    keepLiveMode = true;
                }
            }

            UserCredential credential;

            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "CFCopier.client_secret.json";

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {

                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                LogAnEvent("Credential file saved to: " + credPath);
            }

            if (keepLiveMode)
            {
                return;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });


            #region Loading config file

            //Getting dates
            var culture = new System.Globalization.CultureInfo("ru-RU");
            var targetDay = culture.DateTimeFormat.GetDayName(DateTime.Today.AddDays(1).DayOfWeek);
            var targetDate = DateTime.Today.AddDays(1).Date.ToString(culture.DateTimeFormat.ShortDatePattern);
            var currentDay = culture.DateTimeFormat.GetDayName(DateTime.Today.DayOfWeek);
            var currentDate = DateTime.Today.Date.ToString(culture.DateTimeFormat.ShortDatePattern);

            var currentDayNumber = (int)DateTime.Today.DayOfWeek;
            if (currentDayNumber == 0)
            {
                currentDayNumber = 7;
            }

            LogAnEvent("Reading CFCConfig.xml");
            var configText = System.IO.File.ReadAllText("CFCConfig.xml");
            if (string.IsNullOrWhiteSpace(configText))
            {
                LogAnError("Check CFCConfig.xml file");
                return;
            }

            configText = configText.Replace("$targetDate", targetDate)
                                   .Replace("$targetDay", targetDay)
                                   .Replace("$currentDate", currentDate)
                                   .Replace("$currentDay", currentDay);

            var configXml = XDocument.Parse(configText);

            var spreadsheetIdXElement = configXml.Descendants().FirstOrDefault(l => l.Name.LocalName == "SpreadSheetId");
            if (spreadsheetIdXElement == null)
            {
                LogAnError("There is no SpreadSheetId in CFCConfig.xml file");
                return;
            }

            var spreadsheetId = spreadsheetIdXElement.Value.Trim();

            var rangeXElement = configXml.Descendants().FirstOrDefault(l => l.Name.LocalName == "Range");
            if (rangeXElement == null)
            {
                LogAnError("There is no Range in CFCConfig.xml file");
                return;
            }

            var range = rangeXElement.Value.Trim();

            var daysToWorkWith = configXml.Descendants().Where(l => l.Name.LocalName == "Day"
                                                                    && l.Attribute("numbers") != null
                                                                    && l.Attribute("numbers")
                                                                        .Value.Contains(currentDayNumber.ToString())).ToList();

            if (!daysToWorkWith.Any())
            {
                LogAnError("No work for today!");
                return;
            }

            var tablesNotToHideText =
                daysToWorkWith.Where(l => l.Attribute("tablesNotToHide") != null)
                    .Max(l => l.Attribute("tablesNotToHide").Value);

            int tablesNotToHide = 0;
            int.TryParse(tablesNotToHideText, out tablesNotToHide);

            var datesNotToHide = new List<string>();
            for (int i = 0; i < tablesNotToHide; i++)
            {
                datesNotToHide.Add(DateTime.Today.AddDays(-i).Date.ToString(culture.DateTimeFormat.ShortDatePattern));
            }

            var variablesToReplaceDictionary = new Dictionary<string, string>();
            var variablesToReplaceXElemens = daysToWorkWith.Descendants().Where(l => l.Name.LocalName == "Replace");
            foreach (var replaceElement in variablesToReplaceXElemens)
            {
                var variableName = replaceElement.Attribute("variableName");
                if (variableName != null
                    && !string.IsNullOrWhiteSpace(variableName.Value)
                    && !variablesToReplaceDictionary.ContainsKey(variableName.Value))
                {
                    variablesToReplaceDictionary.Add(variableName.Value, replaceElement.Value.Trim());

                }

            }

            #endregion

            var pagesToHide = new List<int>();

            var spreadSheetReq = service.Spreadsheets.Get(spreadsheetId);
            var spreadSheet = spreadSheetReq.Execute();

            foreach (var sheet in spreadSheet.Sheets)
            {
                if (sheet.Properties.Title.ToUpper() == SourcePageName.ToUpper())
                {
                    if (sheet.Properties.SheetId.HasValue)
                    {
                        SourcePageId = sheet.Properties.SheetId.Value;
                    }

                }

                if (sheet.Properties.SheetId.HasValue && (sheet.Properties.Hidden == null || sheet.Properties.Hidden.HasValue && !sheet.Properties.Hidden.Value))
                {
                    if (!datesNotToHide.Any(sheet.Properties.Title.Contains))
                    {
                        pagesToHide.Add(sheet.Properties.SheetId.Value);
                    }

                }

            }

            if (SourcePageId == -1)
            {
                LogAnError($"Default page named \"{SourcePageName}\" is missing.");
                return;
            }

            #endregion

            #region Creating a new page

            //Creating a new page
            try
            {
                CurrentSpreadSheetName = targetDate + " " + targetDay;
                LogAnEvent($"Creating page \"{CurrentSpreadSheetName}\".");
                var CreateNewPage = new Request()
                {
                    DuplicateSheet = new DuplicateSheetRequest()
                    {
                        SourceSheetId = SourcePageId,
                        NewSheetName = CurrentSpreadSheetName

                    }

                };

                RunBatchRequest(new List<Request>() { CreateNewPage }, service, spreadsheetId);

            }
            catch (Exception)
            {
                LogAnError($"Error! Table \"{CurrentSpreadSheetName}\" already exists.");
                return;

                CurrentSpreadSheetName = targetDate;
                LogAnEvent($"Creating page \"{CurrentSpreadSheetName}\".");
                try
                {
                    var CreateNewPage = new Request()
                    {
                        DuplicateSheet = new DuplicateSheetRequest()
                        {
                            SourceSheetId = SourcePageId,
                            NewSheetName = CurrentSpreadSheetName,

                        }

                    };

                    RunBatchRequest(new List<Request>() { CreateNewPage }, service, spreadsheetId);
                }
                catch (Exception)
                {
                    LogAnError($"Error! Table \"{CurrentSpreadSheetName}\" already exists.");
                    return;
                }

            }

            //getting a new page
            LogAnEvent($"Getting a new table {CurrentSpreadSheetName}.");
            spreadSheet = spreadSheetReq.Execute();
            var currentSpreadSheet =
                spreadSheet.Sheets.FirstOrDefault(l => l.Properties.Title.ToUpper() == CurrentSpreadSheetName.ToUpper());
            if (currentSpreadSheet == null || !currentSpreadSheet.Properties.SheetId.HasValue)
            {
                LogAnError($"Error getting a new table \"{CurrentSpreadSheetName}\".");
                return;
            }

            //showing a new page
            LogAnEvent($"Making \"{CurrentSpreadSheetName}\" visible.");
            var sheetToShowRequest = new Request()
            {
                UpdateSheetProperties = new UpdateSheetPropertiesRequest()
                {
                    Properties = new SheetProperties()
                    {
                        SheetId = currentSpreadSheet.Properties.SheetId.Value,
                        Hidden = false,
                    },

                    Fields = "hidden"
                }

            };

            RunBatchRequest(new List<Request>() { sheetToShowRequest }, service, spreadsheetId);

            #endregion

            #region Getting cells to replace

            var cellsToReplaceList = new List<CellWithText>();

            var spreadsheetsValuesGetRequest = service.Spreadsheets.Values.Get(spreadsheetId, CurrentSpreadSheetName + "!" + range);
            var spreadsheetsValuesGetResponse = spreadsheetsValuesGetRequest.Execute();


            var spreadsheetsValues = spreadsheetsValuesGetResponse.Values;
            if (spreadsheetsValues != null && spreadsheetsValues.Count > 0)
            {
                for (int i = 0; i < spreadsheetsValues.Count; i++)
                {
                    if (spreadsheetsValues[i] != null)
                    {
                        for (int j = 0; j < spreadsheetsValues[i].Count; j++)
                        {
                            if (spreadsheetsValues[i][j] != null)
                            {
                                var stringValue = spreadsheetsValues[i][j].ToString();
                                if (!string.IsNullOrWhiteSpace(stringValue) && variablesToReplaceDictionary.Keys.Any(s => stringValue.Contains("{" + s + "}")))
                                {
                                    foreach (var key in variablesToReplaceDictionary.Keys)
                                    {
                                        var keyToReplace = "{" + key + "}";
                                        if (stringValue.Contains(keyToReplace))
                                        {
                                            cellsToReplaceList.Add(new CellWithText()
                                            {
                                                Column = j,
                                                Row = i,
                                                Text = stringValue.Replace(keyToReplace, variablesToReplaceDictionary[key])
                                            });
                                        }

                                    }

                                }

                            }

                        }

                    }

                }

            }

            #endregion

            #region Hiding active sheets

            //hide active sheets
            LogAnEvent("Hiding other tables.");
            var hideRequestsList = new List<Request>();
            foreach (var sheetToHide in pagesToHide)
            {
                var sheetToHideRequest = new Request()
                {
                    UpdateSheetProperties = new UpdateSheetPropertiesRequest()
                    {
                        Properties = new SheetProperties()
                        {
                            SheetId = sheetToHide,
                            Hidden = true,
                        },

                        Fields = "hidden"
                    }

                };

                hideRequestsList.Add(sheetToHideRequest);
            }

            RunBatchRequest(hideRequestsList, service, spreadsheetId);
            #endregion

            #region Formating cells
            //formating cells
            LogAnEvent("Filling data.");
            var cellTextReplaceRequestsList = new List<Request>();
            foreach (var cellToReplace in cellsToReplaceList)
            {
                cellTextReplaceRequestsList.Add(new Request()
                {
                    UpdateCells = new UpdateCellsRequest()
                    {
                        Start = new GridCoordinate()
                        {
                            SheetId = currentSpreadSheet.Properties.SheetId.Value,
                            ColumnIndex = cellToReplace.Column,
                            RowIndex = cellToReplace.Row
                        },
                        Rows = new List<RowData>()
                        {
                            new RowData()
                            {
                                Values = new List<CellData>()
                                {
                                    new CellData()
                                    {
                                        UserEnteredValue = new ExtendedValue()
                                        {
                                            StringValue = cellToReplace.Text
                                        }
                                    }
                                }

                            }

                        },

                        Fields = "UserEnteredValue"
                    }
                });

            }

            RunBatchRequest(cellTextReplaceRequestsList, service, spreadsheetId);

            ////formating header
            //CellData headerCell = new CellData();

            //headerCell.UserEnteredValue = new ExtendedValue()
            //{
            //    StringValue = string.Format(
            //        "КРОССФИТ {0} {1}. Записаться можно либо на утреннее, либо на вечернее занятие, но не на оба сразу :) Запись до 20:30 {2} {3}",
            //            targetDate, targetDay, currentDate, currentDay)
            //};
            //headerCell.UserEnteredFormat = new CellFormat()
            //{
            //    //BackgroundColor = new Color() {Red = 255, Green = 0, Blue = 0},
            //    TextFormat = new TextFormat() { Bold = true, FontSize = 12, ForegroundColor = new Color() { Alpha = 1, Red = 1, Green = 0, Blue = 0 } }
            //};

            ////formatting morning header
            //CellData headerCellMorning = new CellData();
            //headerCellMorning.UserEnteredValue = new ExtendedValue()
            //{
            //    StringValue = string.Format("8:00 {0} {1}.", targetDate, targetDay)
            //};
            //headerCellMorning.UserEnteredFormat = new CellFormat()
            //{
            //    //BackgroundColor = new Color() { Red = 1, Green = (float?)0.8980392156862746, Blue = (float?)0.603921568627451, Alpha = 1 },
            //    TextFormat = new TextFormat() { Bold = true, FontSize = 12 },
            //    HorizontalAlignment = "center"

            //};


            ////formatting evening header
            //CellData headerCellEvening = new CellData();
            //headerCellEvening.UserEnteredValue = new ExtendedValue()
            //{
            //    StringValue = string.Format("20:30 {0} {1}", targetDate, targetDay)
            //};
            //headerCellEvening.UserEnteredFormat = new CellFormat()
            //{
            //    //BackgroundColor = new Color() { Red = 1, Green = (float?)0.8980392156862746, Blue = (float?)0.603921568627451, Alpha = 1 },
            //    TextFormat = new TextFormat() { Bold = true, FontSize = 12 },
            //    HorizontalAlignment = "center"

            //};

            //var changeHeaderCellRequest = new Request();
            //changeHeaderCellRequest.UpdateCells = new UpdateCellsRequest()
            //{
            //    Start = new GridCoordinate()
            //    {
            //        SheetId = currentSpreadSheet.Properties.SheetId.Value,
            //        ColumnIndex = 0,
            //        RowIndex = 0
            //    },
            //    Rows = new List<RowData>()
            //        {
            //            new RowData()
            //            {
            //                Values = new List<CellData>()
            //                {
            //                    headerCell
            //                }

            //            }
            //        },

            //    Fields = "UserEnteredValue"
            //};

            //var changeHeaderCellMorningRequest = new Request();
            //changeHeaderCellMorningRequest.UpdateCells = new UpdateCellsRequest()
            //{
            //    Start = new GridCoordinate()
            //    {
            //        SheetId = currentSpreadSheet.Properties.SheetId.Value,
            //        ColumnIndex = 0,
            //        RowIndex = 1
            //    },
            //    Rows = new List<RowData>()
            //        {
            //            new RowData()
            //            {
            //                Values = new List<CellData>()
            //                {
            //                    headerCellMorning
            //                }

            //            }
            //        },

            //    Fields = "UserEnteredValue"
            //};

            //var changeHeaderCellEveningRequest = new Request();
            //changeHeaderCellEveningRequest.UpdateCells = new UpdateCellsRequest()
            //{
            //    Start = new GridCoordinate()
            //    {
            //        SheetId = currentSpreadSheet.Properties.SheetId.Value,
            //        ColumnIndex = 5,
            //        RowIndex = 1
            //    },
            //    Rows = new List<RowData>()
            //        {
            //            new RowData()
            //            {
            //                Values = new List<CellData>()
            //                {
            //                    headerCellEvening
            //                }

            //            }
            //        },

            //    Fields = "UserEnteredValue"
            //};

            //RunBatchRequest(new List<Request>() { changeHeaderCellRequest, changeHeaderCellMorningRequest, changeHeaderCellEveningRequest }, service, spreadsheetId);

            #endregion

            #region Sending notifications

            var notificationsXElement = configXml.Descendants().FirstOrDefault(l => l.Name.LocalName == "Notifications");
            if (notificationsXElement.Nodes().Count() > 0)
            {
                foreach (var node in notificationsXElement.Nodes().OfType<XElement>())
                {
                    var serverUrlXElement = node.Elements("ServerUrl").FirstOrDefault();
                    var fromIdXElement = node.Elements("FromId").FirstOrDefault();
                    var recipientIdXElement = node.Elements("RecipientId").FirstOrDefault();
                    var conversationIdXElement = node.Elements("ConversationId").FirstOrDefault();
                    var messageXElement = node.Elements("Message").FirstOrDefault();

                    if (serverUrlXElement != null
                        && fromIdXElement != null
                        && recipientIdXElement != null
                        && conversationIdXElement != null
                        && messageXElement != null)
                    {
                        using (var client = new WebClient())
                        {
                            try
                            {
                                var responseString =
                                client.DownloadString(
                                    $"{serverUrlXElement.Value}?from={fromIdXElement.Value}&recipient={recipientIdXElement.Value}&conversation={conversationIdXElement.Value}&message={messageXElement.Value}");
                            }
                            catch (Exception)
                            {

                            }

                        }

                    }

                }

            }

            #endregion

        }

        private static void RunBatchRequest(List<Request> requests, SheetsService service, string spreadsheetId)
        {
            if (requests.Count == 0)
            {
                return;
            }

            var batchUpdateReq = new BatchUpdateSpreadsheetRequest() { Requests = requests };
            service.Spreadsheets.BatchUpdate(batchUpdateReq, spreadsheetId).Execute();
        }

        private static void LogAnEvent(string message)
        {
            using (EventLog eventLog = new EventLog("CFAppLog", ".", "CFAppLog"))
            {
                //eventLog.Source = "CFCopierApplication";
                eventLog.WriteEntry(message, EventLogEntryType.Information);
            }
        }

        private static void LogAnError(string message)
        {
            using (EventLog eventLog = new EventLog("CFAppLog", ".", "CFAppLog"))
            {
                //eventLog.Source = "CFCopierApplication";
                eventLog.WriteEntry(message, EventLogEntryType.Error);
            }
        }

        public class CellWithText
        {
            public int Row { get; set; }
            public int Column { get; set; }

            public string Text { get; set; }

        }
    }
}