using ExcelTableExportImport;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TablesTest
{
    [TestClass]
    public class TablesTest
    {
        readonly private string inPath = "prices_export_20210812064812.xlsx";
        readonly private string outPath = "prices_export_20210812064812_new.xlsx";


        [TestMethod]
        public void CurrencyEqualsTest()
        {
            ExcelPriceModel[] ImportedModels = Program.ReadFromExcelFile(inPath);
            Program.SaveToExcelFile(ImportedModels, outPath);
            ExcelPriceModel[] ExportedModels = Program.ReadFromExcelFile(outPath);

            Assert.AreEqual(ImportedModels[3].models["Currency"], ExportedModels[3].models["Currency"]);
        }

        [TestMethod]
        public void SkuEqualsTest()
        {
            ExcelPriceModel[] ImportedModels = Program.ReadFromExcelFile(inPath);
            Program.SaveToExcelFile(ImportedModels, outPath);
            ExcelPriceModel[] ExportedModels = Program.ReadFromExcelFile(outPath);

            Assert.AreEqual(ImportedModels[2].models["SKU"], ExportedModels[2].models["SKU"]);
        }

        [TestMethod]
        public void ColumnsNamesCountTest()
        {
            ExcelPriceModel[] ImportedModels = Program.ReadFromExcelFile(inPath);
            Program.SaveToExcelFile(ImportedModels, outPath);
            ExcelPriceModel[] ExportedModels = Program.ReadFromExcelFile(outPath);

            int ImportedModelsCount = 0;
            int ExportedModelsCount = 0;

            foreach (var model in ImportedModels[0].models) ImportedModelsCount++;
            foreach (var model in ExportedModels[0].models) ExportedModelsCount++;

            Assert.AreEqual(ImportedModelsCount, ExportedModelsCount);
        }

        [TestMethod]
        public void ModelsEqualsTest()
        {
            ExcelPriceModel[] ImportedModels = Program.ReadFromExcelFile(inPath);
            Program.SaveToExcelFile(ImportedModels, outPath);
            ExcelPriceModel[] ExportedModels = Program.ReadFromExcelFile(outPath);

            Assert.AreEqual(ImportedModels[0].models.Values.Count, ExportedModels[0].models.Values.Count);
        }
    }
}
