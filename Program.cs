using ClosedXML.Excel;
using System.Collections;

namespace ExcellTableExportImport
{
    public class ExcelPriceModel
    {
        public ArrayList Array { get; set; }

        public ExcelPriceModel(ArrayList array)
        {
            this.Array = array;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // TEST
            string inPath = "C:\\Users\\tuman\\Desktop\\prices_export_20210812064812.xlsx";
            string outPath = "C:\\Users\\tuman\\Desktop\\prices_export_20210812064812_new.xlsx";
            ExcelPriceModel[] res = ReadFromExcelFile(inPath);
            SaveToExcelFile(res, outPath);
        }

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

            ExcelPriceModel[] models = new ExcelPriceModel[columnLength+1];

            for (int i = 1; i <= columnLength; i++)
            {
                var currentRow = worksheet.Row(i);

                ArrayList array = new ArrayList();

                for (int j = 1; j <= rowLength; j++)
                {
                    array.Add(currentRow.Cell(j));
                }

                models[i] = new ExcelPriceModel(array);

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
                    foreach (var data in model.Array)
                    {
                        worksheet.Cell(row, column).Value = data;
                        column++;
                    }
                    row++;
                    column = 1;
                }
            }
            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(path);
        }
    }
}
