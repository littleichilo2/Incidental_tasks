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
        List<string> musicfile_list = new List<string>(){
            "open-i.mp3","practice1.mp3","pi1bgquestion.mp3","question1choice1.mp3","attention-i.mp3","story.mp3","pi-text1.mp3",
            "question2choice2.mp3","pquestion-i11.mp3","pquestion-i12.mp3","pquestion-i13.mp3","practice 2.mp3","pi2bgquestion.mp3",
            "pi-text2.mp3","pquestion-i21.mp3","pquestion-i22.mp3","pquestion-i23.mp3","pquestion-i24.mp3","Istart1.mp3","ibgquestion1.mp3",
            "question1choice1.mp3","attention-i.mp3","story.mp3","i-text1.mp3","question2choice2.mp3","question-i11.mp3","question-i12.mp3",
            "question-i13.mp3","question-i14.mp3","istart2.mp3","ibgquestion2.mp3","question1choice1.mp3","attention-i.mp3","story.mp3",
            "i-text2.mp3","question2choice2.mp3","question-i21.mp3","question-i22.mp3","question-i23.mp3","Istart3.mp3","Ibgquestion3.mp3",
            "question1choice1.mp3","attention-i.mp3","story.mp3","i-text3.mp3","question2choice2.mp3","question-i31.mp3","question-i32.mp3",
            "Istart4.mp3","ibgquestion1.mp3","question1choice1.mp3","attention-i.mp3","story.mp3","i-text4.mp3","question2choice2.mp3","question-i41.mp3",
            "question-i42.mp3","question-i43.mp3","question-i44.mp3","question-i45.mp3","Istart5.mp3","Ibgquestion5.mp3","question1choice1.mp3","attention-i.mp3",
            "story.mp3","i-text5.mp3","question2choice2.mp3","question-i51.mp3","question-i52.mp3","question-i53.mp3","question-i54.mp3","question-i55.mp3",
            "Istart6.mp3","Ibgquestion6.mp3","question1choice1.mp3","attention-i.mp3","story.mp3","i-text6.mp3","question2choice2.mp3",
            "question-i61.mp3","question-i62.mp3","question-i63.mp3","question-i64.mp3","question-i65.mp3","Istart7.mp3","Ibgquestion7.mp3",
            "question1choice1.mp3","attention-i.mp3","story.mp3","i-text7.mp3","question2choice2.mp3","question-i71.mp3","question-i72.mp3",
            "question-i73.mp3","question-i74.mp3","question-i75.mp3"
        };
        List<int> music_list = new List<int>();

        List<string> answerlist = new List<string>();
        List<double> answer_timelist = new List<double>();
        List<int> againtimes_list = new List<int>();
        int again = 0;
        DateTime datetime_now;
        DateTime datetimeflag;
        WMPLib.WindowsMediaPlayer musicplayer = new WMPLib.WindowsMediaPlayer();
        int frame;
        public Form1()
        {
            InitializeComponent();
            entranceButton.Click += delegate (object sender, EventArgs e) { button_Click(sender, e, MessageType.entranceButton); };
            flowerbutton.Click += delegate (object sender, EventArgs e) { button_Click(sender, e, MessageType.flowerbutton); };
            leafbutton.Click += delegate (object sender, EventArgs e) { button_Click(sender, e, MessageType.leafbutton); };
            circlebutton.Click += delegate (object sender, EventArgs e) { button_Click(sender, e, MessageType.circlebutton); };
            Xbutton.Click += delegate (object sender, EventArgs e) { button_Click(sender, e, MessageType.Xbutton); };
            Button1.Click += delegate (object sender, EventArgs e) { button_Click(sender, e, MessageType.Button1); };
            Button2.Click += delegate (object sender, EventArgs e) { button_Click(sender, e, MessageType.Button2); };
            Button3.Click += delegate (object sender, EventArgs e) { button_Click(sender, e, MessageType.Button3); };
            nextbutton.Click += delegate (object sender, EventArgs e) { button_Click(sender, e, MessageType.nextbutton); };


        }

        private void button_Click(object sender, EventArgs e, MessageType type)
        {
            //add the againtimes into list before the next frame
            if (again_buttonlist.IndexOf(frame) != -1)
                againtimes_list.Add(again);
            again = 0;

            frame = frame + 1;
            datetimeflag = DateTime.Now;
            //music play
            //musicplayer.URL = music_list[frame];
            musicplayer.controls.play();

            //What is going to appear on the panel
            if (titlelist.IndexOf(frame) != -1)
            {
                titlelabel.Visible = true;
            }
            if (flowerlist.IndexOf(frame) != -1)
            {
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
            if (type.Equals(MessageType.circlebutton))
            {
                datetime_now = DateTime.Now;
                answer_timelist.Add((datetime_now - datetimeflag).TotalMilliseconds);
                answerlist.Add("O");
            }
            else if (type.Equals(MessageType.Xbutton))
            {
                datetime_now = DateTime.Now;
                answer_timelist.Add((datetime_now - datetimeflag).TotalMilliseconds);
                answerlist.Add("X");
            }
            else if (type.Equals(MessageType.Button1))
            {
                datetime_now = DateTime.Now;
                answer_timelist.Add((datetime_now - datetimeflag).TotalMilliseconds);
                answerlist.Add("1");
            }
            else if (type.Equals(MessageType.Button2))
            {
                datetime_now = DateTime.Now;
                answer_timelist.Add((datetime_now - datetimeflag).TotalMilliseconds);
                answerlist.Add("2");
            }
            else if (type.Equals(MessageType.Button3))
            {
                datetime_now = DateTime.Now;
                answer_timelist.Add((datetime_now - datetimeflag).TotalMilliseconds);
                answerlist.Add("3");
            }


        }

        private void againbutton_click(object sender, EventArgs e)
        {
            again += 1;
            //musicplayer.URL = music_list[frame];
            musicplayer.controls.play();
        }



        private void endbutton_click(object sender, EventArgs e)
        {
            string nowpath = System.IO.Path.GetDirectoryName(Application.ExecutablePath);

            string pathFile = nowpath +"\\推論理解";

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
                excelApp.Cells[1, 1] = "流水號";
                excelApp.Cells[1, 2] = "名稱";
                excelApp.Cells[1, 3] = "出生年月";
                excelApp.Cells[1, 4] = "施測日期";
                excelApp.Cells[1, 5] = "性別";
                excelApp.Cells[1, 6] = "第一題";
                excelApp.Cells[1, 7] = "作答時間";
                excelApp.Cells[1, 8] = "重複次數";
                excelApp.Cells[1, 9] = "第二題";
                excelApp.Cells[1, 10] = "作答時間";
                excelApp.Cells[1, 11] = "重複次數";
                excelApp.Cells[1, 12] = "第三題";
                excelApp.Cells[1, 13] = "作答時間";
                excelApp.Cells[1, 14] = "重複次數";
                excelApp.Cells[1, 15] = "第四題";
                excelApp.Cells[1, 16] = "作答時間";
                excelApp.Cells[1, 17] = "重複次數";

                excelApp.Cells[1, 18] = "第五題";
                excelApp.Cells[1, 19] = "作答時間";
                excelApp.Cells[1, 20] = "重複次數";

                excelApp.Cells[1, 21] = "第六題";
                excelApp.Cells[1, 22] = "作答時間";
                excelApp.Cells[1, 23] = "重複次數";

                excelApp.Cells[1, 24] = "第七題";
                excelApp.Cells[1, 25] = "作答時間";
                excelApp.Cells[1, 26] = "重複次數";

                excelApp.Cells[1, 27] = "第八題";
                excelApp.Cells[1, 28] = "作答時間";
                excelApp.Cells[1, 29] = "重複次數";

                excelApp.Cells[1, 30] = "第九題";
                excelApp.Cells[1, 31] = "作答時間";
                excelApp.Cells[1, 32] = "重複次數";

                excelApp.Cells[1, 33] = "第十題";
                excelApp.Cells[1, 34] = "作答時間";
                excelApp.Cells[1, 35] = "重複次數";


                // 設定第1列顏色
                wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 2]];
                wRange.Select();
                wRange.Font.Color = ColorTranslator.ToOle(Color.White);
                wRange.Interior.Color = ColorTranslator.ToOle(Color.DimGray);

                // 設定第2列資料
                /*
                excelApp.Cells[2, 1] = ;//流水號
                excelApp.Cells[2, 2] = ;//姓名
                excelApp.Cells[2, 3] = ;//出生年月
                excelApp.Cells[2, 4] = ;//施測年月
                excelApp.Cells[2, 5] = ;//性別*/
                excelApp.Cells[2, 6] = answerlist[0];
                excelApp.Cells[2, 7] = answer_timelist[0];
                excelApp.Cells[2, 8] = againtimes_list[0];
                excelApp.Cells[2, 9] = answerlist[1];
                excelApp.Cells[2, 10] = answer_timelist[1];
                excelApp.Cells[2, 11] = againtimes_list[1];
                excelApp.Cells[2, 12] = answerlist[2];
                excelApp.Cells[2, 13] = answer_timelist[2];
                excelApp.Cells[2, 14] = againtimes_list[2];
                excelApp.Cells[2, 15] = answerlist[3];
                excelApp.Cells[2, 16] = answer_timelist[3];
                excelApp.Cells[2, 17] = againtimes_list[3];
                excelApp.Cells[2, 18] = answerlist[4];
                excelApp.Cells[2, 19] = answer_timelist[4];
                excelApp.Cells[2, 20] = againtimes_list[4];
                excelApp.Cells[2, 21] = answerlist[5];
                excelApp.Cells[2, 22] = answer_timelist[5];
                excelApp.Cells[2, 23] = againtimes_list[5];
                excelApp.Cells[2, 24] = answerlist[6];
                excelApp.Cells[2, 25] = answer_timelist[6];
                excelApp.Cells[2, 26] = againtimes_list[6];
                excelApp.Cells[2, 27] = answerlist[7];
                excelApp.Cells[2, 28] = answer_timelist[7];
                excelApp.Cells[2, 29] = againtimes_list[7];
                excelApp.Cells[2, 30] = answerlist[8];
                excelApp.Cells[2, 31] = answer_timelist[8];
                excelApp.Cells[2, 32] = againtimes_list[8];
                excelApp.Cells[2, 33] = answerlist[9];
                excelApp.Cells[2, 34] = answer_timelist[9];
                excelApp.Cells[2, 35] = againtimes_list[9];

                // 設定第3列資料
                //excelApp.Cells[3, 1] = "BB";
                //excelApp.Cells[3, 2] = "20";

                // 設定第4列資料
                //excelApp.Cells[4, 1] = "CC";
                //excelApp.Cells[4, 2] = "30";

                // 設定第5列資料
                //excelApp.Cells[5, 1] = "總計";
                // 設定總和公式 =SUM(B2:B4)
                //excelApp.Cells[5, 2].Formula = string.Format("=SUM(B{0}:B{1})", 2, 4);
                // 設定第5列顏色
                //wRange = wSheet.Range[wSheet.Cells[5, 1], wSheet.Cells[5, 2]];
                //wRange.Select();
                //wRange.Font.Color = ColorTranslator.ToOle(Color.Red);
                //wRange.Interior.Color = ColorTranslator.ToOle(Color.Yellow);

                // 自動調整欄寬
                //wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[5, 2]];
                //wRange.Select();
                //wRange.Columns.AutoFit();

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
        enum MessageType
        {
            entranceButton,
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
