using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json;
using System;

namespace GoogleSheetAPI
{
    internal class Copy
    {
        public SheetProperties CopySheets(string ClientId, string ClientSecret, string destinationSpreadsheetId, string sourceSpreadsheetId, int sourceSheetId)
        {
            var sheetsService = GoogleAPI.GetSheetsService(ClientId, ClientSecret);

            CopySheetToAnotherSpreadsheetRequest requestBody = new CopySheetToAnotherSpreadsheetRequest();
            requestBody.DestinationSpreadsheetId = destinationSpreadsheetId;

            SpreadsheetsResource.SheetsResource.CopyToRequest request = sheetsService.Spreadsheets.Sheets.CopyTo(requestBody, sourceSpreadsheetId, sourceSheetId);

            // To execute asynchronously in an async method, replace `request.Execute()` as shown:
            SheetProperties response = request.Execute();
            // Data.SheetProperties response = await request.ExecuteAsync();

            // TODO: Change code below to process the `response` object:
            Console.WriteLine(JsonConvert.SerializeObject(response));

            return response;
        }
    }
}