using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;


namespace Inference
{
    public partial class Form1 : Form
    {
        List<int> titlelist = new List<int>();
        List<int> flowerlist = new List<int>();
        List<int> leaflist = new List<int>();
        List<int> circlelist = new List<int>();
        List<int> Xlist = new List<int>();
        List<int> list_1 = new List<int>();
        List<int> list_2 = new List<int>();
        List<int> list_3 = new List<int>();
        List<int> again_buttonlist = new List<int>();
        List<int> next_buttonlist = new List<int>();
        List<string> music_list = new List<string>();
        List<string> answerlist = new List<string>();
        List<string> answer_timelist = new List<string>();
        DateTime datetime_now;
        DateTime datetimeflag;
        WMPLib.WindowsMediaPlayer musicplayer = new WMPLib.WindowsMediaPlayer();
        int frame;
        public Form1()
        {
            InitializeComponent();
            entrancebutton.Click += delegate(object sender, EventArgs e) { button_Click(sender, e, MessageType.entrancebutton); };  
            flowerbutton.Click += delegate(object sender, EventArgs e) { button_Click(sender, e,  MessageType.flowerbutton); };  
            leafbutton.Click += delegate(object sender, EventArgs e) { button_Click(sender, e,  MessageType.leafbutton); };  
            circlebutton.Click += delegate(object sender, EventArgs e) { button_Click(sender, e,  MessageType.circlebutton); };  
            Xbutton.Click += delegate(object sender, EventArgs e) { button_Click(sender, e,  MessageType.Xbutton); };  
            Button1.Click += delegate(object sender, EventArgs e) { button_Click(sender, e,  MessageType.Button1); };  
            Button2.Click += delegate(object sender, EventArgs e) { button_Click(sender, e,  MessageType.Button2); };  
            Button3.Click += delegate(object sender, EventArgs e) { button_Click(sender, e,  MessageType.Button3); };  
            nextbutton.Click += delegate(object sender, EventArgs e) { button_Click(sender, e,  MessageType.nextbutton); };  

            
        }

        private void button_Click(object sender, EventArgs e, MessageType type)
        {
            frame = frame + 1;
            datetimeflag = DateTime.Now;
            //music play
            musicplayer.URL = music_list[frame];
            musicplayer.controls.play();

            //What is going to appear on the panel
            if (titlelist.IndexOf(frame) != -1)
            {
                titlelabel.Visible = true;
            }
            if (flowerlist.IndexOf(frame) != -1){
                flowerbutton.Visible = true;
            }
            if (leaflist.IndexOf(frame) != -1)
            {
                leafbutton.Visible = true;
            }
            if (circlelist.IndexOf(frame) != -1)
            {
                circlebutton.Visible = true;
            }
            if (Xlist.IndexOf(frame) != -1)
            {
                Xbutton.Visible = true;
            }
            if (list_1.IndexOf(frame) != -1)
            {
                Button1.Visible = true;
            }
            if (list_2.IndexOf(frame) != -1)
            {
                Button2.Visible = true;
            }
            if (list_3.IndexOf(frame) != -1)
            {
                Button3.Visible = true;
            }
            if (again_buttonlist.IndexOf(frame) != -1)
            {
                againbutton.Visible = true;
            }
            if (next_buttonlist.IndexOf(frame) != -1)
            {
                nextbutton.Visible = true;
            }
            
            //Record the response from the button
            if(type.Equals(MessageType.circlebutton)){
                datetime_now = DateTime.Now;
                answer_timelist.Add((datetime_now - datetimeflag).TotalMilliseconds);
                answerlist.Add('Yes');
            }
            else if(type.Equals(MessageType.Xbutton)){
                datetime_now = DateTime.Now;
                answer_timelist.Add((datetime_now - datetimeflag).TotalMilliseconds);
                answerlist.Add('No');
            }
            else if(type.Equals(MessageType.Button1)){
                datetime_now = DateTime.Now;
                answer_timelist.Add((datetime_now - datetimeflag).TotalMilliseconds);
                answerlist.Add('1');
            }
            else if(type.Equals(MessageType.Button2)){
                datetime_now = DateTime.Now;
                answer_timelist.Add((datetime_now - datetimeflag).TotalMilliseconds);
                answerlist.Add('2');
            }
            else if(type.Equals(MessageType.Button3)){
                datetime_now = DateTime.Now;
                answer_timelist.Add((datetime_now - datetimeflag).TotalMilliseconds);
                answerlist.Add('3');
            }

            
        }
        
        private void againbutton_click(object sender,EventArgs e){
            musicplayer.URL = music_list[frame];
            musicplayer.controls.play();
        }


        
        private void endbutton_click(object sender, EventArgs e){
            string nowpath = System.IO.Path.GetDirectoryName(Application.ExecutablePath);

            string pathFile = nowpath+'推論理解';
         
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
                wSheet.Name = "sheet1";
         
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
                    //Console.WriteLine("儲存文件於 " + Environment.NewLine + pathFile);
                }
                catch (Exception ex)
                {
                    //Console.WriteLine("儲存檔案出錯，檔案可能正在使用" + Environment.NewLine + ex.Message);
                }
            }
            catch (Exception ex)
            {
                //Console.WriteLine("產生表時出錯！" + Environment.NewLine + ex.Message);
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
         
            //Console.Read();
        }
        enum MessageType{
            entrancebutton,
            flowerbutton,
            leafbutton,
            circlebutton,
            Xbutton,
            Button1,
            Button2,
            Button3,
            nextbutton
        }
    }
}
