flowerframe
leafframe
circleframe 
Xframe
musicframe



if button click:
  datetime2 = datetime.now()
  responsetime = datetime2 - datetime1 
  musicframe[frame].play
  frame++
  if framelist[frame] in targetframe:
    target.Enable = True
  else:
    target.Enable = False

private void endbutton_click(){
  Excel.Application xlApp = new Excel.Application();
  Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"test.xlsx");
  Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
  Excel.Range xlRange = xlWorksheet.UsedRange;





  GC.Collect();
  GC.WaitForPendingFinalizers();

  Marshal.ReleaseComObject(xlRange);
  Marshal.ReleaseComObject(xlWorksheet);

  xlWorkbook.Close();
  Marshal.ReleaseComObject(xlWorkbook);

  xlApp.Quit();
  Marshal.ReleaseComObject(xlApp);
}



send arguments in button click
public Form1()  
{  
  InitializeComponent();  
  
  button1.Click += delegate(object sender, EventArgs e) { button_Click(sender, e, "This is   From Button1", MessageType.B1); };  
  
  button2.Click += delegate(object sender, EventArgs e) { button_Click(sender, e, "This is From Button2", MessageType.B2); };  
  
}  
  
void button_Click(object sender, EventArgs e, string message, MessageType type)  
{  
   if (type.Equals(MessageType.B1))  
   {  
      label1.Text = message;  
   }  
   else if (type.Equals(MessageType.B2))  
   {  
      label1.Text = message;  
   }  
}  
  
enum MessageType  
{  
   B1,  
   B2  
}  





//another example
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

static void Main(string[] args)
{
    // 設定儲存檔名，不用設定副檔名，系統自動判斷 excel 版本，產生 .xls 或 .xlsx 副檔名
    string pathFile = @"D:\test";
 
    Excel.Application excelApp;
    Excel._Workbook wBook;
    Excel._Worksheet wSheet;
    Excel.Range wRange;
 
    // 開啟一個新的應用程式
    excelApp = new Excel.Application();
 
    // 讓Excel文件可見
    excelApp.Visible = true;
 
    // 停用警告訊息
    excelApp.DisplayAlerts = false;
 
    // 加入新的活頁簿
    excelApp.Workbooks.Add(Type.Missing);
 
    // 引用第一個活頁簿
    wBook = excelApp.Workbooks[1];
 
    // 設定活頁簿焦點
    wBook.Activate();
 
    try
    {
        // 引用第一個工作表
        wSheet = (Excel._Worksheet)wBook.Worksheets[1];
 
        // 命名工作表的名稱
        wSheet.Name = "工作表測試";
 
        // 設定工作表焦點
        wSheet.Activate();
 
        excelApp.Cells[1, 1] = "Excel測試";
 
        // 設定第1列資料
        excelApp.Cells[1, 1] = "名稱";
        excelApp.Cells[1, 2] = "數量";
        // 設定第1列顏色
        wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 2]];
        wRange.Select();
        wRange.Font.Color = ColorTranslator.ToOle(Color.White);
        wRange.Interior.Color = ColorTranslator.ToOle(Color.DimGray);
 
        // 設定第2列資料
        excelApp.Cells[2, 1] = "AA";
        excelApp.Cells[2, 2] = "10";
 
        // 設定第3列資料
        excelApp.Cells[3, 1] = "BB";
        excelApp.Cells[3, 2] = "20";
 
        // 設定第4列資料
        excelApp.Cells[4, 1] = "CC";
        excelApp.Cells[4, 2] = "30";
 
        // 設定第5列資料
        excelApp.Cells[5, 1] = "總計";
        // 設定總和公式 =SUM(B2:B4)
        excelApp.Cells[5, 2].Formula = string.Format("=SUM(B{0}:B{1})", 2, 4);
        // 設定第5列顏色
        wRange = wSheet.Range[wSheet.Cells[5, 1], wSheet.Cells[5, 2]];
        wRange.Select();
        wRange.Font.Color = ColorTranslator.ToOle(Color.Red);
        wRange.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
 
        // 自動調整欄寬
        wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[5, 2]];
        wRange.Select();
        wRange.Columns.AutoFit();
 
        try
        {
            //另存活頁簿
            wBook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Console.WriteLine("儲存文件於 " + Environment.NewLine + pathFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine("儲存檔案出錯，檔案可能正在使用" + Environment.NewLine + ex.Message);
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine("產生報表時出錯！" + Environment.NewLine + ex.Message);
    }
 
    //關閉活頁簿
    wBook.Close(false, Type.Missing, Type.Missing);
 
    //關閉Excel
    excelApp.Quit();
 
    //釋放Excel資源
    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
    wBook = null;
    wSheet = null;
    wRange = null;
    excelApp = null;
    GC.Collect();
 
    Console.Read();
}
