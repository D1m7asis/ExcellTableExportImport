using ClosedXML.Excel;
using System;
using System.Collections.Generic;

namespace ExcelTableExportImport
{
    public class ExcelPriceModel
    {
        public Dictionary<string, object> models;

        public ExcelPriceModel(Dictionary<string, object> models)
        {
            this.models = models;
        }
    }

    public class Program
    {
        static void Main() { }

        public static ExcelPriceModel[] ReadFromExcelFile(string path)
        {
            var workbook = new XLWorkbook(path);
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RangeUsed().RowsUsed();

            int columnLength = 0;
            int rowLength = 0;

            foreach (var columnSizeCheck in worksheet.RowsUsed())
            {
                columnLength++;
            }

            foreach (var rowSizeCheck in worksheet.ColumnsUsed())
            {
                rowLength++;
            }

            ExcelPriceModel[] models = new ExcelPriceModel[columnLength + 1];

            for (int i = 0; i < columnLength; i++)
            {
                var currentRow = worksheet.Row(i + 1);

                Dictionary<string, object> dict = new Dictionary<string, object>();

                for (int j = 0; j < rowLength; j++)
                {
                    string columnName = worksheet.Row(1).Cell(j + 1).Value.ToString();
                    string cellData = currentRow.Cell(j + 1).Value.ToString();

                    switch (columnName)
                    {
                        case "Sale price":
                        case "Min quantity":
                        case "List price":
                            if (!string.IsNullOrEmpty(cellData))
                            {
                                if (!cellData.Equals(columnName) & int.TryParse(cellData, out int value))
                                {
                                    dict.Add(columnName, value);
                                }

                                else dict.Add(columnName, cellData);
                            }
                            else dict.Add(columnName, "");

                            break;

                        case "Modified":
                        case "Valid from":
                        case "Valid to":
                        case "Created date":
                            if (!String.IsNullOrEmpty(cellData))
                            {
                                if (!cellData.Equals(columnName))
                                {
                                    string[] dateNtime = cellData.Split(" ");
                                    string[] data = dateNtime[0].Split(".");
                                    string[] time = dateNtime[1].Split(":");
                                    DateTime dt = new DateTime(int.Parse(data[2]), int.Parse(data[1]), int.Parse(data[0]), int.Parse(time[0]), int.Parse(time[1]), 0);
                                    dict.Add(columnName, dt);
                                }
                                else
                                {
                                    dict.Add(columnName, cellData);
                                }
                            }
                            else dict.Add(columnName, "");
                            break;

                        default:
                            try
                            {
                                dict.Add(columnName, cellData);
                            }
                            catch (System.ArgumentException) { }

                            break;
                    }
                }
                models[i] = new ExcelPriceModel(dict);
            }
            return models;
        }
        public static void SaveToExcelFile(ExcelPriceModel[] priceModels, string path)
        {
            IXLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add("VirtoCommerce");

            int row = 1;    // ряд
            int column = 1; // столбец

            foreach (var model in priceModels)
            {
                if (model != null)
                {
                    foreach (var data in model.models)
                    {
                        if ((data.Key.Equals("Modified") |
                           data.Key.Equals("Valid from") | 
                           data.Key.Equals("Valid to") | 
                           data.Key.Equals("Created date")) 
                           & !data.Key.Equals(data.Value))
                        {
                            if (DateTime.TryParse(data.Value.ToString(), out DateTime dt))
                            {
                                worksheet.Cell(row, column).Value = $"{dt.Day}.{dt.Month}.{dt.Year} {dt:HH:mm}";
                            }
                            else
                            {
                                worksheet.Cell(row, column).Value = data.Value;
                            }
                        }

                        else if ((data.Key.Equals("List price") | 
                            data.Key.Equals("Sale price") | 
                            data.Key.Equals("Min quantity"))
                            & !data.Key.Equals(data.Value))
                        {
                            if (decimal.TryParse(data.Value.ToString(), out decimal value))
                            {
                                worksheet.Cell(row, column).SetValue(value);
                            }
                        }
                        else
                        {
                            worksheet.Cell(row, column).Value = data.Value;
                        }

                        column++;
                    }
                    row++;
                    column = 1;
                }
            }

            var firstCell = worksheet.FirstCellUsed();
            var lastCell = worksheet.LastCellUsed();
            var range = worksheet.Range(firstCell.Address, lastCell.Address);
            var table = range.CreateTable();

            table.Theme = XLTableTheme.TableStyleLight13;

            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(path);
        }
    }
}