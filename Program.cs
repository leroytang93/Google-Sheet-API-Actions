using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using System.Data;
using System.Threading;
using System.Configuration;
using System;

namespace GoogleSheetAPI
{
    public class GoogleAPI
    {
        public static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        public static string ClientId = ConfigurationManager.AppSettings.Get("ClientID");
        public static string ClientSecret = ConfigurationManager.AppSettings.Get("ClientSecret");
        public static string sourceSpreadsheetId = ConfigurationManager.AppSettings.Get("sourceSpreadsheetId");
        public static string sourceSheetId = ConfigurationManager.AppSettings.Get("sourceSheetId");
        public static string destinationSpreadsheetId = ConfigurationManager.AppSettings.Get("destinationSpreadsheetId");
        public static bool Success;

        public static void Main(string[] args)
        {
            // Add New Sheet
            Add_Sheet add = new Add_Sheet();
            //add.AddSheet(ClientId, ClientSecret, sourceSpreadsheetId, "New_Sheet_Added");

            // Get all Sheets
            Get_Sheet_Properties sheetProperties = new Get_Sheet_Properties();
            // Retrieves the first instance of specific column name within desired range in spreadsheet
            //sheetProperties.RenameSheet(ClientId, ClientSecret, destinationSpreadsheetId, 439585049, "RENAME TEST1", out Success);

            int colIndex = sheetProperties.GetColumnIndex(ClientId, ClientSecret, destinationSpreadsheetId, "Form Responses 1", "Screen Date & Time", out Success);

            int rowIndexbasedon2Strings = sheetProperties.GetRowIndexbasedonTwoStrings(ClientId, ClientSecret, destinationSpreadsheetId, "Sheet1!A:G", "column 3", "44", "column 5", "6", out Success);

            int rowIndex = sheetProperties.GetRowIndex(ClientId, ClientSecret, destinationSpreadsheetId, "Sheet1!A:G", "44", out Success);

            // Retrieves the first instance of specified value in row of desired range in spreadsheet
            DataTable result = sheetProperties.GetAllSheets(ClientId, ClientSecret, sourceSpreadsheetId, out Success);

            Delete delete = new Delete();
            // Delete specific Sheet from Spreadsheet
            delete.DeleteSheet(ClientId, ClientSecret, destinationSpreadsheetId, 1505537265);
                // Delete specific Row from Spreadsheet
                delete.DeleteRow(ClientId, ClientSecret, sourceSpreadsheetId, Int32.Parse(sourceSheetId), 2);

                Copy copy = new Copy();
            // Copy sheets from Spreadsheet
            copy.CopySheets(ClientId, ClientSecret, destinationSpreadsheetId, sourceSpreadsheetId, Int32.Parse(sourceSheetId));
        }
        public static SheetsService GetSheetsService(string clientId, string clientSecret)
        {
            var clientSecrets = new ClientSecrets
            {
                ClientId = clientId,
                ClientSecret = clientSecret
            };

            return new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer =
                    GoogleWebAuthorizationBroker.AuthorizeAsync(clientSecrets, Scopes, "admin", CancellationToken.None).Result
            });
        }

        public static DriveService GetDriveService(string clientId, string clientSecret)
        {
            var clientSecrets = new ClientSecrets
            {
                ClientId = clientId,
                ClientSecret = clientSecret,
            };

            return new DriveService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GoogleWebAuthorizationBroker.AuthorizeAsync(clientSecrets, new[] { DriveService.Scope.DriveFile }, "admin", CancellationToken.None).Result
            });
        }
    }
}

        
