using System;
using NPOI;
using NPOI.HSSF.UserModel;
//using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

namespace esOO
{
    public class TesOO
    {
        public static  void doit()
        {
            // Open Template
            FileStream fs = new FileStream((@"D:\templ.xls"), FileMode.Open, FileAccess.Read);

            // Load the template into a NPOI workbook
            HSSFWorkbook templateWorkbook = new HSSFWorkbook(fs, true);

            // Load the sheet you are going to use as a template into NPOI
            HSSFSheet sheet =(HSSFSheet) templateWorkbook.GetSheet("工作表1");
            for(int i=1;i<45;i++)
            {
                 for(int j = 2; j < 10; j++)
                {
                    sheet.GetRow(i).GetCell(j).SetCellValue(i * 100 + j);
                }
            }
            //Row.CreateCell(0).CellFormula = "SUM(A1:B1)";
            //Row.CreateCell(1).CellFormula = "A1-B1";
            //更新有公式的欄位
            sheet.ForceFormulaRecalculation = true;
            FileStream file = new FileStream(@"d:\tmp_npoi.xls", FileMode.Create);//產生檔案
            templateWorkbook.Write(file);
            fs.Close();
            file.Close();
        }
    }
}
