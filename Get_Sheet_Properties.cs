using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace GoogleSheetAPI
{
    internal class Get_Sheet_Properties
    {
        public DataTable GetAllSheets(string ClientID, string ClientSecret, string spreadSheetID, out bool Success)
        {
            GoogleAPI google = new GoogleAPI();
            var table = new DataTable();
            DataColumn SheetTitle = table.Columns.Add("Sheet Title");
            SheetTitle.DataType = System.Type.GetType("System.String");
            DataColumn SheetID = table.Columns.Add("Sheet ID");
            SheetID.DataType = System.Type.GetType("System.Int32");
            try
            {
                var sheetsService = GoogleAPI.GetSheetsService(ClientID,ClientSecret);

                var results = sheetsService.Spreadsheets.Get(spreadSheetID).Execute();

                for (int row = 0; row < results.Sheets.Count; row++)
                {
                    DataRow dataRow = table.NewRow();
                    dataRow["Sheet ID"] = results.Sheets[row].Properties.SheetId;
                    dataRow["Sheet Title"] = results.Sheets[row].Properties.Title;
                    table.Rows.Add(dataRow);
                }
                Success = true;
                return table;
            }
            catch
            {
                Success = false;
                return table;
            }
        }

        public int GetColumnIndex(string ClientID, string ClientSecret, string spreadSheetID, string range, string searchString, out bool Success)
        {
            var sheetsService = GoogleAPI.GetSheetsService(ClientID, ClientSecret);
            int colIndex = 0;
            // Post request to get range of values in spreadsheet
            var rangeRequest = sheetsService.Spreadsheets.Values.Get(spreadSheetID, range);
            var cellValues = rangeRequest.Execute().Values;
            var collection = cellValues.ToList();

            if (cellValues == null)
            {
                Success = false;
                return 0;
            }
            else
            {
                // remove all empty records from List
                // just need to loop through cellValues[0] for column index
                foreach (var row in cellValues)
                {
                    if (row.Count == 0)
                    {
                        collection.Remove(row);
                    }
                }
                for (colIndex = 0; colIndex <= collection[0].Count; colIndex++)
                {
                    if (collection[0][colIndex].ToString() == searchString)
                    {
                        break;
                    }
                }
                Success = true;
                return colIndex;
            }
        }

        public int GetRowIndex(string ClientID, string ClientSecret, string spreadSheetID, string range, string searchString, out bool Success)
        {
            var sheetsService = GoogleAPI.GetSheetsService(ClientID, ClientSecret);
            int rowIndex = 0;
            // Post request to get range of values in spreadsheet
            var rangeRequest = sheetsService.Spreadsheets.Values.Get(spreadSheetID, range);
            var cellValues = rangeRequest.Execute().Values;
            var collection = cellValues.ToList();
            Success = false;

            if (cellValues == null)
            {
                Success = false;
                return 0;
            }
            else
            {
                // remove all empty records from List
                // just need to loop through cellValues[0] for column index
                foreach (var row in cellValues)
                {
                    if (row.Count == 0)
                    {
                        collection.Remove(row);
                    }
                }
                foreach (var row in collection)
                {
                    for (int cellReference = 0; cellReference < row.Count; cellReference++)
                    {
                        if (row[cellReference].ToString() == searchString)
                        {
                            Success = true;
                            return rowIndex;
                        }
                    }
                    rowIndex++;
                }
                Success = false;
                return 0;
            }
        }

        public int GetRowIndexbasedonTwoStrings(string ClientID, string ClientSecret, string spreadSheetID, string range, string columnName1, string searchString1, string columnName2, string searchString2, out bool Success)
        {
            var sheetsService = GoogleAPI.GetSheetsService(ClientID, ClientSecret);
            int rowIndex = 0;

            // Post request to get range of values in spreadsheet
            var rangeRequest = sheetsService.Spreadsheets.Values.Get(spreadSheetID, range);
            var cellValues = rangeRequest.Execute().Values;
            var collection = cellValues.ToList();

            Success = false;

            int col1Index = GetColumnIndex(ClientID, ClientSecret, spreadSheetID, range, columnName1, out bool col1Success);
            int col2Index = GetColumnIndex(ClientID, ClientSecret, spreadSheetID, range, columnName2, out bool col2Success);

            if (cellValues == null)
            {
                Success = false;
                return 0;
            }
            else
            {
                // remove all empty records from List
                // just need to loop through cellValues[0] for column index
                foreach (var row in cellValues)
                {
                    if (row.Count == 0)
                    {
                        collection.Remove(row);
                    }
                }
                foreach (var row in collection)
                {
                    if (row[col1Index].ToString() == searchString1 && row[col2Index].ToString() == searchString2)
                    {
                        Success = true;
                        return rowIndex;
                    }
                    rowIndex++;
                }
                Success = false;
                return 0;
            }
        }

        public void RenameSheet(string ClientId, string ClientSecret, string spreadSheetID, int sheetID, string newName, out bool Success)
        {
            try
            {
                var sheetsService = GoogleAPI.GetSheetsService(ClientId, ClientSecret);

                var request = new Request
                {
                    UpdateSheetProperties = new UpdateSheetPropertiesRequest
                    {
                        Properties = new SheetProperties()
                        {
                            SheetId = sheetID,
                            Title = newName
                        },
                        Fields = "Title"
                    }
                };

                var RenameSheetRequest = new BatchUpdateSpreadsheetRequest { Requests = new List<Request> { request } };

                sheetsService.Spreadsheets.BatchUpdate(RenameSheetRequest, spreadSheetID).Execute();

                Success = true;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                Success = false;
            }
        }

        public void RenameSpreadsheet(string ClientId, string ClientSecret, string spreadSheetID, string newName, out bool Success)
        {
            try
            {
                var sheetsService = GoogleAPI.GetSheetsService(ClientId, ClientSecret);

                var request = new Request
                {
                    UpdateSpreadsheetProperties = new UpdateSpreadsheetPropertiesRequest
                    {
                        Properties = new SpreadsheetProperties()
                        {
                            Title = newName
                        },
                        Fields = "Title"
                    }
                };

                var RenameSheetRequest = new BatchUpdateSpreadsheetRequest { Requests = new List<Request> { request } };

                sheetsService.Spreadsheets.BatchUpdate(RenameSheetRequest, spreadSheetID).Execute();

                Success = true;
            }
            catch
            {
                Success = false;
            }
        }
    }
}