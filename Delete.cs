using Google.Apis.Sheets.v4.Data;
using System.Collections.Generic;

namespace GoogleSheetAPI
{
    internal class Delete
    {
        public void DeleteSheet(string ClientId, string ClientSecret, string spreadSheetID, int sheetID)
        {
            var sheetsService = GoogleAPI.GetSheetsService(ClientId, ClientSecret);

            var request = new Request
            {
                DeleteSheet = new DeleteSheetRequest
                {
                    SheetId = sheetID
                }
            };

            var deletesheetRequest = new BatchUpdateSpreadsheetRequest { Requests = new List<Request> { request } };

            //execute request
            sheetsService.Spreadsheets.BatchUpdate(deletesheetRequest, spreadSheetID).Execute();
        }

        public void DeleteRow(string ClientId, string ClientSecret, string spreadSheetID, int sheetID, int rowIndex)
        {
            var sheetsService = GoogleAPI.GetSheetsService(ClientId, ClientSecret);

            var request = new Request
            {
                DeleteDimension = new DeleteDimensionRequest
                {
                    Range = new DimensionRange
                    {
                        SheetId = sheetID,
                        Dimension = "ROWS",
                        StartIndex = rowIndex,
                        EndIndex = rowIndex + 1
                    }
                }
            };

            var deleteRowRequest = new BatchUpdateSpreadsheetRequest { Requests = new List<Request> { request } };

            sheetsService.Spreadsheets.BatchUpdate(deleteRowRequest, spreadSheetID).Execute();
        }

        public void DeleteColumn(string ClientId, string ClientSecret, string spreadSheetID, int sheetID, int colIndex)
        {
            var sheetsService = GoogleAPI.GetSheetsService(ClientId, ClientSecret);

            var request = new Request
            {
                DeleteDimension = new DeleteDimensionRequest
                {
                    Range = new DimensionRange
                    {
                        SheetId = sheetID,
                        Dimension = "COLUMNS",
                        StartIndex = colIndex,
                        EndIndex = colIndex + 1
                    }
                }
            };

            var deleteColRequest = new BatchUpdateSpreadsheetRequest { Requests = new List<Request> { request } };

            sheetsService.Spreadsheets.BatchUpdate(deleteColRequest, spreadSheetID).Execute();
        }
    }
}