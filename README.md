"" 
"# esOO" 

名詞說明:
HSSF － 提供讀寫Microsoft Excel XLS格式檔案的功能。
XSSF － 提供讀寫Microsoft Excel OOXML XLSX格式檔案的功能。
HWPF － 提供讀寫Microsoft Word DOC格式檔案的功能。
HSLF － 提供讀寫Microsoft PowerPoint格式檔案的功能。
HDGF － 提供讀Microsoft Visio格式檔案的功能。
HPBF － 提供讀Microsoft Publisher格式檔案的功能。
HSMF － 提供讀Microsoft Outlook格式檔案的功能。 
有關POI的說明可參考:
http://en.wikipedia.org/wiki/Apache_POI
NPOI軟體下載網址:
http://npoi.codeplex.com/
本範例使用版本為npoi 2.0.1 (beta1)
軟體解壓縮後，將\npoi 2.0.1 (beta1)\npoi\dotnet3.5的DLL檔複製到自己專案的bin目錄下
加入參考: NPOI.dll, NPOI.OOXML.dll
加入命名空間:
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
程式碼:
//範例一，簡單產生Excel檔案的方法
private void CreateExcelFile()
{
    //建立Excel 2003檔案
    IWorkbook wb = new HSSFWorkbook();
    ISheet ws = wb.CreateSheet("Class");
 
    ////建立Excel 2007檔案
    //IWorkbook wb = new XSSFWorkbook();
    //ISheet ws = wb.CreateSheet("Class");
 
    ws.CreateRow(0);//第一行為欄位名稱
    ws.GetRow(0).CreateCell(0).SetCellValue("name");
    ws.GetRow(0).CreateCell(1).SetCellValue("score");
    ws.CreateRow(1);//第二行之後為資料
    ws.GetRow(1).CreateCell(0).SetCellValue("abey");
    ws.GetRow(1).CreateCell(1).SetCellValue(85);
    ws.CreateRow(2);
    ws.GetRow(2).CreateCell(0).SetCellValue("tina");
    ws.GetRow(2).CreateCell(1).SetCellValue(82);
    ws.CreateRow(3);
    ws.GetRow(3).CreateCell(0).SetCellValue("boi");
    ws.GetRow(3).CreateCell(1).SetCellValue(84);
    ws.CreateRow(4);
    ws.GetRow(4).CreateCell(0).SetCellValue("hebe");
    ws.GetRow(4).CreateCell(1).SetCellValue(86);
    ws.CreateRow(5);
    ws.GetRow(5).CreateCell(0).SetCellValue("paul");
    ws.GetRow(5).CreateCell(1).SetCellValue(82);
    FileStream file = new FileStream(@"d:\tmp\npoi.xls", FileMode.Create);//產生檔案
    wb.Write(file);
    file.Close();
}
 
//範例二，DataTable轉成Excel檔案的方法
private void DataTableToExcelFile(DataTable dt)
{
    //建立Excel 2003檔案
    IWorkbook wb = new HSSFWorkbook();
    ISheet ws;
 
    ////建立Excel 2007檔案
    //IWorkbook wb = new XSSFWorkbook();
    //ISheet ws;
 
    if (dt.TableName != string.Empty)
    {
        ws = wb.CreateSheet(dt.TableName);
    }
    else
    {
        ws = wb.CreateSheet("Sheet1");
    }
 
    ws.CreateRow(0);//第一行為欄位名稱
    for (int i = 0; i < dt.Columns.Count; i++)
    {
         ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
    }
 
    for (int i = 0; i < dt.Rows.Count; i++)
    {
        ws.CreateRow(i + 1);
        for (int j = 0; j < dt.Columns.Count; j++)
        {
            ws.GetRow(i+1).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
        }
    }
     
    FileStream file = new FileStream(@"d:\tmp\npoi.xls", FileMode.Create);//產生檔案
    wb.Write(file);
    file.Close();
}

Use NPOI to populate an Excel template
So I’ve been using NPOI all week and decide to do a quick “demo” for my team today. My demo was to show how to use NPOI to populate (update) an Excel template that includes various charts.  Even though NPOI does not support creating charts from scratch, it does support updating files that already include (hence template) charts. I started by going to the Microsoft website where they have a bunch of free “pretty” templates on, randomly choosing one with a bunch of formulas and charts.  It took me about an hour to build the complete demo using very simple and easy to read code. Most of the updates or just putting values in cells, but the actual process of opening/reading/inserting/saving a new or existing file in NPOI is very easy for a novice programmers.
Development Summary ( Step-by-Step )
Get a template, I grabbed mine from here ( office.microsoft.com ).
** I only used the first sheet and deleted all the sample data
Create a new ASP.NET 3.5 Web Application projecct
Download NPOI binaries and include in your project
Build a UI that will be used to populate your template.
** This could also be populated by a data sources (db, XML, etc..)
** NOTE:  I used Excel to create the form using Excel formulas
Add some c# code “magic” to load data into the template using NPOI
Sounds simple because it is… Here is the code to add the form contents into the Excel template.
?
// Open Template
FileStream fs = new FileStream(Server.MapPath(@"\template\Template_EventBudget.xls"), FileMode.Open, FileAccess.Read);
 
// Load the template into a NPOI workbook
HSSFWorkbook templateWorkbook = new HSSFWorkbook(fs, true);
 
// Load the sheet you are going to use as a template into NPOI
HSSFSheet sheet = templateWorkbook.GetSheet("Event Budget");
 
// Insert data into template
sheet.GetRow(1).GetCell(1).SetCellValue(EventName.Value);  // Inserting a string value into Excel
sheet.GetRow(1).GetCell(5).SetCellValue(DateTime.Parse(EventDate.Value));  // Inserting a date value into Excel
 
sheet.GetRow(5).GetCell(2).SetCellValue(Double.Parse(Roomandhallfees.Value));  // Inserting a number value into Excel
sheet.GetRow(6).GetCell(2).SetCellValue(Double.Parse(Sitestaff.Value));
sheet.GetRow(7).GetCell(2).SetCellValue(Double.Parse(Equipment.Value));
sheet.GetRow(8).GetCell(2).SetCellValue(Double.Parse(Tablesandchairs.Value));
sheet.GetRow(12).GetCell(2).SetCellValue(Double.Parse(Flowers.Value));
sheet.GetRow(13).GetCell(2).SetCellValue(Double.Parse(Candles.Value));
sheet.GetRow(14).GetCell(2).SetCellValue(Double.Parse(Lighting.Value));
sheet.GetRow(15).GetCell(2).SetCellValue(Double.Parse(Balloons.Value));
sheet.GetRow(16).GetCell(2).SetCellValue(Double.Parse(Papersupplies.Value));
sheet.GetRow(20).GetCell(2).SetCellValue(Double.Parse(Graphicswork.Value));
sheet.GetRow(21).GetCell(2).SetCellValue(Double.Parse(Photocopying_Printing.Value));
sheet.GetRow(22).GetCell(2).SetCellValue(Double.Parse(Postage.Value));
sheet.GetRow(26).GetCell(2).SetCellValue(Double.Parse(Telephone.Value));
sheet.GetRow(27).GetCell(2).SetCellValue(Double.Parse(Transportation.Value));
sheet.GetRow(28).GetCell(2).SetCellValue(Double.Parse(Stationerysupplies.Value));
sheet.GetRow(29).GetCell(2).SetCellValue(Double.Parse(Faxservices.Value));
sheet.GetRow(33).GetCell(2).SetCellValue(Double.Parse(Food.Value));
sheet.GetRow(34).GetCell(2).SetCellValue(Double.Parse(Drinks.Value));
sheet.GetRow(35).GetCell(2).SetCellValue(Double.Parse(Linens.Value));
sheet.GetRow(36).GetCell(2).SetCellValue(Double.Parse(Staffandgratuities.Value));
sheet.GetRow(40).GetCell(2).SetCellValue(Double.Parse(Performers.Value));
sheet.GetRow(41).GetCell(2).SetCellValue(Double.Parse(Speakers.Value));
sheet.GetRow(42).GetCell(2).SetCellValue(Double.Parse(Travel.Value));
sheet.GetRow(43).GetCell(2).SetCellValue(Double.Parse(Hotel.Value));
sheet.GetRow(44).GetCell(2).SetCellValue(Double.Parse(Other.Value));
sheet.GetRow(48).GetCell(2).SetCellValue(Double.Parse(Ribbons_Plaques_Trophies.Value));
sheet.GetRow(49).GetCell(2).SetCellValue(Double.Parse(Gifts.Value));
 
// Force formulas to update with new data we added
sheet.ForceFormulaRecalculation = true;
 
// Save the NPOI workbook into a memory stream to be sent to the browser, could have saved to disk.
MemoryStream ms = new MemoryStream();
templateWorkbook.Write(ms);
 ac
// Send the memory stream to the browser
ExportDataTableToExcel(ms, "EventExpenseReport.xls");
//http://www.zachhunter.com/2010/05/npoi-excel-template/

https://stackoverflow.com/questions/5855813/how-to-read-file-using-npoi