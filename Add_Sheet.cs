using Google.Apis.Sheets.v4.Data;
using System.Collections.Generic;

namespace GoogleSheetAPI
{
    internal class Add_Sheet
    {
        public void AddSheet(string ClientId, string ClientSecret, string spreadSheetID, string sheetName)
        {
            var sheetsService = GoogleAPI.GetSheetsService(ClientId, ClientSecret);

            // Add new Sheet
            var addSheetRequest = new AddSheetRequest();
            addSheetRequest.Properties = new SheetProperties();
            addSheetRequest.Properties.Title = sheetName;
            BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateSpreadsheetRequest.Requests = new List<Request>();
            batchUpdateSpreadsheetRequest.Requests.Add(new Request
            {
                AddSheet = addSheetRequest
            });

            var batchUpdateRequest =
                sheetsService.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadSheetID);

            batchUpdateRequest.Execute();
        }
    }
}