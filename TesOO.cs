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
        public static void readxls()
        {
            string filepath = @"d:\tmp_npoi.xls";

            HSSFWorkbook hssfwb;

            using (FileStream file = new FileStream(filepath, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheetAt(0);
           
            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {
                    // Set new cell value
                    //sheet.GetRow(row).GetCell(0).SetCellValue("foo");
                    // Console.WriteLine("Row {0} = {1}", row, sheet.GetRow(row).GetCell(0).StringCellValue);
                    //https://stackoverflow.com/questions/5855813/how-to-read-file-using-npoi
                    for (int col=0;col< sheet.GetRow(row).Cells.Count; col++)
                    {
                        var cell = sheet.GetRow(row).GetCell(col);


                        if (cell != null)
                            switch (cell.CellType)
                            {
                                case CellType.Numeric:
                                    Console.Write(cell.NumericCellValue);
      
                                    break;
                                case CellType.String:
                                    Console.Write(cell.StringCellValue);
                                    break;
                                case CellType.Blank:
                                    Console.Write(string.Empty);
                                   
                                    break;
                                case CellType.Formula:
                                    Console.Write(cell.NumericCellValue);
                                    break;
                            }
                        Console.Write("\t");
                    }
                    Console.WriteLine();
                }
            }
            
            // Save the file
           // using (FileStream file = new FileStream(filepath, FileMode.Open, FileAccess.Write))
            //{
           //     hssfwb.Write(file);
           // }

            Console.ReadLine();

        }
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
