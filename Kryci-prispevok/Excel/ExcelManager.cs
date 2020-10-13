

using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System.IO;
using System.Security.Cryptography;

namespace Kryci_prispevok
{
    class ExcelManager
    {
        public void Spracuj_Excel(string ExcelName)
        {
            if (ExcelName == null)
                return;
            FileInfo Excel = new FileInfo(ExcelName);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var pckg = new ExcelPackage(Excel))
            {
                var exlwk = pckg.Workbook.Worksheets[0];
                var ExcelText = exlwk.Cells["A1"].Value;

                ExcelPackage ExcelPkg = new ExcelPackage();
                ExcelWorksheet wsSheet1 = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
                wsSheet1.Cells["A1"].Value = ExcelText + "something";

                wsSheet1.Protection.IsProtected = false;
                wsSheet1.Protection.AllowSelectLockedCells = false;
                ExcelPkg.SaveAs(new FileInfo("NewExcel/Excel.xlsx"));
            }
           
        }
    }
}
