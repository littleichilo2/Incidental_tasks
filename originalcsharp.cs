using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;



namespace Listening_comprehension
{
    public partial class Form1 : Form
    {
        List<string> namelist = new List<string>();
        List<string> birthlist = new List<string>();
        List<string> testlist = new List<string>();
        List<string> answerList = new List<string>();
        string number;
        string gender;
        TimeSpan totaltime = TimeSpan.FromMilliseconds(0);
        List<int> titlelist = new List<int>(){
            1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,
            30,31,32,33,34,35,36,37,38,39
        };
        List<int> list_1 = new List<int>() { 2, 3, 5, 6, 7, 8, 10, 11, 12, 13, 15, 16, 17, 19, 21, 22, 23, 25, 26, 27, 29, 30, 31, 33, 34, 35, 37, 38, 39 };
        List<int> list_2 = new List<int>() { 2, 3, 5, 6, 7, 8, 10, 11, 12, 13, 15, 16, 17, 19, 21, 22, 23, 25, 26, 27, 29, 30, 31, 33, 34, 35, 37, 38, 39 };
        List<int> list_3 = new List<int>() { 2, 3, 5, 6, 7, 8, 10, 11, 12, 13, 15, 16, 17, 19, 21, 22, 23, 25, 26, 27, 29, 30, 31, 33, 34, 35, 37, 38, 39 };
        List<int> again_buttonlist = new List<int>(){
            2,3,5,6,7,8,10,11,12,13,15,16,17,19,21,22,23,25,26,27,29,30,31,33,34,35,37,38,39
        };
        List<int> next_buttonlist = new List<int>(){
            2,3,5,6,7,8,10,11,12,13,15,16,17,19,21,22,23,25,26,27,29,30,31,33,34,35,37,38,39
        };
        List<int> next_buttonlist2 = new List<int>(){
          1,4,9,14,18,20,24,28,32,36
        };
        List<string> musicfile_list = new List<string>(){
            "1open-lc.mp3","2guide-lc.mp3","3plc-text1.mp3",
            "4plc-choice1.mp3","5plc-answer1.mp3",
            "6plc-choice2.mp3","7plc-answer2.mp3",
            "8Istart1.mp3","2guide-lc.mp3","9lc-text11.mp3",
            "10lc-choice11.mp3",
            "11lc-choice12.mp3",
            "12lc-choice13.mp3",
            "13lc-choice14.mp3",
            "14lc-text12.mp3",
            "15lc-choice21.mp3",
            "16lc-choice22.mp3",
            "17lc-choice23.mp3",
            "18lc-choice24.mp3",
            "19lc-text13.mp3",
            "20lc-choice31.mp3",
            "21lc-choice32.mp3",
            "22lc-choice33.mp3",
            "23open-dic.mp3","24guide-dic2.mp3","25pdic-text21.mp3",
            "26pdic-choice21.mp3","27pdic-choice22.mp3",
            "28Istart2.mp3","29dic-text11.mp3",
            "30dic-choice11.mp3",
            "31dic-choice12.mp3",
            "32dic-choice13.mp3",
            "33dic-text12.mp3",
            "34dic-choice21.mp3",
            "35dic-choice22.mp3",
            "36dic-choice23.mp3",
            "37dic-text13.mp3",
            "38dic-choice31.mp3",
            "39dic-choice32.mp3",
            "40dic-choice33.mp3",
            "41dic-text14.mp3",
            "42dic-choice41.mp3",
            "43dic-choice42.mp3",
            "44dic-choice43.mp3",
            "45dic-text15.mp3",
            "46dic-choice51.mp3",
            "47dic-choice52.mp3",
            "48dic-choice53.mp3"
        };
        List<List<int>> music_list = new List<List<int>>();


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
            nextbutton.Click += delegate (object sender, EventArgs e) { button_Click(sender, e, MessageType.nextbutton); };
            music_list.Add(new List<int> { 0 });
            music_list.Add(new List<int> { 0, 1, 2 });
            music_list.Add(new List<int> { 3, 4 });
            music_list.Add(new List<int> { 5, 6 });
            music_list.Add(new List<int> { 7, 8, 9 });
            music_list.Add(new List<int> { 10 });
            music_list.Add(new List<int> { 11 });
            music_list.Add(new List<int> { 12 });
            music_list.Add(new List<int> { 13 });
            music_list.Add(new List<int> { 14 });
            music_list.Add(new List<int> { 15 });
            music_list.Add(new List<int> { 16 });
            music_list.Add(new List<int> { 17 });
            music_list.Add(new List<int> { 18 });
            music_list.Add(new List<int> { 19 });
            music_list.Add(new List<int> { 20 });
            music_list.Add(new List<int> { 21 });
            music_list.Add(new List<int> { 22 });

            music_list.Add(new List<int> { 23, 24, 25 });
            music_list.Add(new List<int> { 26, 27 });
            music_list.Add(new List<int> { 28, 24, 29 });
            music_list.Add(new List<int> { 30 });
            music_list.Add(new List<int> { 31 });
            music_list.Add(new List<int> { 32 });
            music_list.Add(new List<int> { 33 });
            music_list.Add(new List<int> { 34 });
            music_list.Add(new List<int> { 35 });
            music_list.Add(new List<int> { 36 });
            music_list.Add(new List<int> { 37 });
            music_list.Add(new List<int> { 38 });
            music_list.Add(new List<int> { 39 });
            music_list.Add(new List<int> { 40 });
            music_list.Add(new List<int> { 41 });
            music_list.Add(new List<int> { 42 });
            music_list.Add(new List<int> { 43 });
            music_list.Add(new List<int> { 44 });
            music_list.Add(new List<int> { 45 });
            music_list.Add(new List<int> { 46 });
            music_list.Add(new List<int> { 47 });
            music_list.Add(new List<int> { 48 });
        }

        private void button_Click(object sender, EventArgs e, MessageType type)
        {

            //add the againtimes into list before the next frame
            if (again_buttonlist.IndexOf(frame) != -1)
                againtimes_list.Add(again);
            again = 0;
            titlelabel.Visible = false;

            Button1.Visible = false;
            Button2.Visible = false;
            Button3.Visible = false;
            Button1.Enabled = true;
            Button2.Enabled = true;
            Button3.Enabled = true;

            if (Button1.Checked == true){
                Button1.Checked = false;
            }
            else if(Button2.Checked == true){
                Button2.Checked = false;
            }
            else if (Button3.Checked ==true){
                Button3.Checked = false;
            }


            
            
            againbutton.Visible = false;
            nextbutton.Visible = false;

            frame = frame + 1;
            Console.WriteLine("frame:" + frame);

            //music play
            WMPLib.IWMPPlaylist playlist = musicplayer.playlistCollection.newPlaylist("playlist");
            //axWindowsMediaPlayer1
            totaltime = TimeSpan.FromMilliseconds(0);
            if (frame < 39)
            {
                for (int k = 0; k < music_list[frame].Count(); k++)
                {

                    WMPLib.IWMPMedia media = musicplayer.newMedia(musicfile_list[music_list[frame][k]]);
                    totaltime = totaltime + TimeSpan.FromMilliseconds(media.duration);
                    playlist.appendItem(media);
                }
                musicplayer.currentPlaylist = playlist;
                musicplayer.settings.setMode("shuffle", false);
                musicplayer.controls.play();
            }
            //What is going to appear on the panel
            if (titlelist.IndexOf(frame) != -1)
            {
                titlelabel.Visible = true;
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
                nextbutton.Text = "下一題";
            }
            if (next_buttonlist2.IndexOf(frame) != -1)
            {
                nextbutton.Text = "開始答題";
                nextbutton.Visible = true;
            }
            if (frame == 40)
            {
                endbutton.Visible = true;
            }

            //Record the response from the button

            if (type.Equals(MessageType.entranceButton))
            {
                teachingpanel.Visible = true;
                datetimeflag = DateTime.Now;
                number = textBox3.Text;
                namelist.Add(textBox1.Text);
                namelist.Add(textBox2.Text);
                birthlist.Add(dateTimePicker1.Text);
                testlist.Add(dateTimePicker2.Text);
                if (radioButton1.Checked)
                {
                    gender = "男";
                }
                else if (radioButton2.Checked)
                {
                    gender = "女";
                }
            }
            else if (type.Equals(MessageType.nextbutton))
            {
                datetimeflag = DateTime.Now;
            }

        }

        private void Button1_CheckedChanged(object sender, EventArgs e)
        {
            Button1.Enabled = false;
            Button2.Enabled = false;
            Button3.Enabled = false;
            datetime_now = DateTime.Now;
            answer_timelist.Add((datetime_now - datetimeflag - totaltime).TotalMilliseconds);
            answerlist.Add("1");
            datetimeflag = DateTime.Now;
        }
        private void Button2_CheckedChanged(object sender, EventArgs e)
        {
            Button1.Enabled = false;
            Button2.Enabled = false;
            Button3.Enabled = false;
            datetime_now = DateTime.Now;
            answer_timelist.Add((datetime_now - datetimeflag - totaltime).TotalMilliseconds);
            answerlist.Add("2");
            datetimeflag = DateTime.Now;
        }
        private void Button3_CheckedChanged(object sender, EventArgs e)
        {
            Button1.Enabled = false;
            Button2.Enabled = false;
            Button3.Enabled = false;
            datetime_now = DateTime.Now;
            answer_timelist.Add((datetime_now - datetimeflag - totaltime).TotalMilliseconds);
            answerlist.Add("3");
            datetimeflag = DateTime.Now;
        }


        private void againbutton_click(object sender, EventArgs e)
        {
            again += 1;
            WMPLib.IWMPPlaylist playlist = musicplayer.playlistCollection.newPlaylist("playlist");
            //axWindowsMediaPlayer1
            for (int k = 0; k < music_list[frame].Count(); k++)
            {
                WMPLib.IWMPMedia media = musicplayer.newMedia(musicfile_list[music_list[frame][k]]);
                Console.WriteLine(media.duration);
                playlist.appendItem(media);
            }
            musicplayer.currentPlaylist = playlist;
            musicplayer.settings.setMode("shuffle", false);
            musicplayer.controls.play();
        }



        private void endbutton_click(object sender, EventArgs e)
        {
            string nowpath = System.IO.Path.GetDirectoryName(Application.ExecutablePath);

            string pathFile = nowpath + "\\推論理解";// + namelist[0] + namelist[1];


            Excel.Application excelApp = new Excel.Application();
            string[] Files = Directory.GetFiles(nowpath);
            if (Files.IndexOf("推論理解") != -1)
            {
                Excel.Workbook wBook = excelApp.Workbooks.Open(pathFile);
                Excel.Worksheet wSheet = wBook.Sheets["推論理解"];
                Excel.Range wRange = wSheet.UsedRange;
                excelApp.Cells[wRange.Rows.Count + 1, 1] = number;//流水號
                excelApp.Cells[wRange.Rows.Count + 1, 2] = namelist[0] + namelist[1];//姓名
                excelApp.Cells[wRange.Rows.Count + 1, 3] = birthlist[0];//出生年月
                excelApp.Cells[wRange.Rows.Count + 1, 4] = testlist[0];//施測年月
                excelApp.Cells[wRange.Rows.Count + 1, 5] = gender;//性別
                if (answerlist[0] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 6] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 6] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 7] = answer_timelist[0];
                excelApp.Cells[wRange.Rows.Count + 1, 8] = againtimes_list[0];
                if (answerlist[1] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 9] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 9] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 10] = answer_timelist[1];
                excelApp.Cells[wRange.Rows.Count + 1, 11] = againtimes_list[1];
                if (answerlist[wRange.Rows.Count + 1] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 12] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 12] = "0";
                }

                excelApp.Cells[wRange.Rows.Count + 1, 13] = answer_timelist[wRange.Rows.Count + 1];
                excelApp.Cells[wRange.Rows.Count + 1, 14] = againtimes_list[wRange.Rows.Count + 1];
                if (answerlist[3] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 15] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 15] = "0";
                }

                excelApp.Cells[wRange.Rows.Count + 1, 16] = answer_timelist[3];
                excelApp.Cells[wRange.Rows.Count + 1, 17] = againtimes_list[3];
                if (answerlist[4] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 18] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 18] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 19] = answer_timelist[4];
                excelApp.Cells[wRange.Rows.Count + 1, 20] = againtimes_list[4];
                if (answerlist[5] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 21] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 21] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 22] = answer_timelist[5];
                excelApp.Cells[wRange.Rows.Count + 1, 23] = againtimes_list[5];
                if (answerlist[6] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 24] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 24] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 25] = answer_timelist[6];
                excelApp.Cells[wRange.Rows.Count + 1, 26] = againtimes_list[6];
                if (answerlist[7] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 27] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 27] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 28] = answer_timelist[7];
                excelApp.Cells[wRange.Rows.Count + 1, 29] = againtimes_list[7];
                if (answerlist[8] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 30] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 30] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 31] = answer_timelist[8];
                excelApp.Cells[wRange.Rows.Count + 1, 32] = againtimes_list[8];
                if (answerlist[9] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 33] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 33] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 34] = answer_timelist[9];
                excelApp.Cells[wRange.Rows.Count + 1, 35] = againtimes_list[9];
                if (answerlist[10] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 36] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 36] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 37] = answer_timelist[10];
                excelApp.Cells[wRange.Rows.Count + 1, 38] = againtimes_list[10];
                if (answerlist[11] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 39] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 39] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 40] = answer_timelist[11];
                excelApp.Cells[wRange.Rows.Count + 1, 41] = againtimes_list[11];
                if (answerlist[12] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 42] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 42] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 43] = answer_timelist[12];
                excelApp.Cells[wRange.Rows.Count + 1, 44] = againtimes_list[12];
                if (answerlist[13] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 45] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 45] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 46] = answer_timelist[13];
                excelApp.Cells[wRange.Rows.Count + 1, 47] = againtimes_list[13];
                if (answerlist[14] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 48] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 48] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 49] = answer_timelist[14];
                excelApp.Cells[wRange.Rows.Count + 1, 50] = againtimes_list[14];
                if (answerlist[15] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 51] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 51] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 52] = answer_timelist[15];
                excelApp.Cells[wRange.Rows.Count + 1, 53] = againtimes_list[15];
                if (answerlist[16] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 54] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 54] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 55] = answer_timelist[16];
                excelApp.Cells[wRange.Rows.Count + 1, 56] = againtimes_list[16];
                if (answerlist[17] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 57] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 57] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 58] = answer_timelist[17];
                excelApp.Cells[wRange.Rows.Count + 1, 59] = againtimes_list[17];
                if (answerlist[18] == "3")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 60] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 60] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 61] = answer_timelist[18];
                excelApp.Cells[wRange.Rows.Count + 1, 62] = againtimes_list[18];
                if (answerlist[19] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 63] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 63] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 64] = answer_timelist[19];
                excelApp.Cells[wRange.Rows.Count + 1, 65] = againtimes_list[19];
                if (answerlist[20] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 66] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 66] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 67] = answer_timelist[20];
                excelApp.Cells[wRange.Rows.Count + 1, 68] = againtimes_list[20];
                if (answerlist[21] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 69] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 69] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 70] = answer_timelist[21];
                excelApp.Cells[wRange.Rows.Count + 1, 71] = againtimes_list[21];
                if (answerlist[22] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 72] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 72] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 73] = answer_timelist[22];
                excelApp.Cells[wRange.Rows.Count + 1, 74] = againtimes_list[22];
                if (answerlist[23] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 75] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 75] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 76] = answer_timelist[23];
                excelApp.Cells[wRange.Rows.Count + 1, 77] = againtimes_list[23];
                if (answerlist[24] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 78] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 78] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 79] = answer_timelist[24];
                excelApp.Cells[wRange.Rows.Count + 1, 80] = againtimes_list[24];
                if (answerlist[25] == "")
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 81] = "1";
                }
                else
                {
                    excelApp.Cells[wRange.Rows.Count + 1, 81] = "0";
                }
                excelApp.Cells[wRange.Rows.Count + 1, 82] = answer_timelist[25];
                excelApp.Cells[wRange.Rows.Count + 1, 83] = againtimes_list[25];
                wBook.Save();
                excelApp.Quit();

                //釋放Excel資源
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                wBook = null;
                wSheet = null;
                //wRange = null;
                excelApp = null;
                GC.Collect();
                Application.Exit();
                //Console.Read();
            }
            else
            {
                Excel.Workbook wBook = excelApp.Workbooks.Add();
                Excel.Worksheet wSheet = new Excel.Worksheet();

                wSheet = wBook.Worksheets[1];
                wSheet.Name = "推論理解";
                try
                {

                    //excelApp.Cells[1, 1] = "Excel測試";

                    // 設定第1列資料
                    excelApp.Cells[1, 1] = "流水號";
                    excelApp.Cells[1, 2] = "姓名";
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

                    excelApp.Cells[1, 36] = "第十一題";
                    excelApp.Cells[1, 37] = "作答時間";
                    excelApp.Cells[1, 38] = "重複次數";
                    excelApp.Cells[1, 39] = "第十二題";
                    excelApp.Cells[1, 40] = "作答時間";
                    excelApp.Cells[1, 41] = "重複次數";
                    excelApp.Cells[1, 42] = "第十三題";
                    excelApp.Cells[1, 43] = "作答時間";
                    excelApp.Cells[1, 44] = "重複次數";
                    excelApp.Cells[1, 45] = "第十四題";
                    excelApp.Cells[1, 46] = "作答時間";
                    excelApp.Cells[1, 47] = "重複次數";
                    excelApp.Cells[1, 48] = "第十五題";
                    excelApp.Cells[1, 49] = "作答時間";
                    excelApp.Cells[1, 50] = "重複次數";
                    excelApp.Cells[1, 51] = "第十六題";
                    excelApp.Cells[1, 52] = "作答時間";
                    excelApp.Cells[1, 53] = "重複次數";
                    excelApp.Cells[1, 54] = "第十七題";
                    excelApp.Cells[1, 55] = "作答時間";
                    excelApp.Cells[1, 56] = "重複次數";
                    excelApp.Cells[1, 57] = "第十八題";
                    excelApp.Cells[1, 58] = "作答時間";
                    excelApp.Cells[1, 59] = "重複次數";
                    excelApp.Cells[1, 60] = "第十九題";
                    excelApp.Cells[1, 61] = "作答時間";
                    excelApp.Cells[1, 62] = "重複次數";
                    excelApp.Cells[1, 63] = "第二十題";
                    excelApp.Cells[1, 64] = "作答時間";
                    excelApp.Cells[1, 65] = "重複次數";
                    excelApp.Cells[1, 66] = "第二十一題";
                    excelApp.Cells[1, 67] = "作答時間";
                    excelApp.Cells[1, 68] = "重複次數";
                    excelApp.Cells[1, 69] = "第二十二題";
                    excelApp.Cells[1, 70] = "作答時間";
                    excelApp.Cells[1, 71] = "重複次數";
                    excelApp.Cells[1, 72] = "第二十三題";
                    excelApp.Cells[1, 73] = "作答時間";
                    excelApp.Cells[1, 74] = "重複次數";
                    excelApp.Cells[1, 75] = "第二十四題";
                    excelApp.Cells[1, 76] = "作答時間";
                    excelApp.Cells[1, 77] = "重複次數";
                    excelApp.Cells[1, 78] = "第二十五題";
                    excelApp.Cells[1, 79] = "作答時間";
                    excelApp.Cells[1, 80] = "重複次數";
                    excelApp.Cells[1, 81] = "第二十六題";
                    excelApp.Cells[1, 82] = "作答時間";
                    excelApp.Cells[1, 83] = "重複次數";



                    // 設定第1列顏色
                    /*wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 2]];
                    wRange.Select();
                    wRange.Font.Color = ColorTranslator.ToOle(Color.White);
                    wRange.Interior.Color = ColorTranslator.ToOle(Color.DimGray);*/

                    // 設定第2列資料

                    excelApp.Cells[2, 1] = number;//流水號
                    excelApp.Cells[2, 2] = namelist[0] + namelist[1];//姓名
                    excelApp.Cells[2, 3] = birthlist[0];//出生年月
                    excelApp.Cells[2, 4] = testlist[0];//施測年月
                    excelApp.Cells[2, 5] = gender;//性別
                    if (answerlist[0] == "")
                    {
                        excelApp.Cells[2, 6] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 6] = "0";
                    }
                    excelApp.Cells[2, 7] = answer_timelist[0];
                    excelApp.Cells[2, 8] = againtimes_list[0];
                    if (answerlist[1] == "")
                    {
                        excelApp.Cells[2, 9] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 9] = "0";
                    }
                    excelApp.Cells[2, 10] = answer_timelist[1];
                    excelApp.Cells[2, 11] = againtimes_list[1];
                    if (answerlist[2] == "")
                    {
                        excelApp.Cells[2, 12] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 12] = "0";
                    }

                    excelApp.Cells[2, 13] = answer_timelist[2];
                    excelApp.Cells[2, 14] = againtimes_list[2];
                    if (answerlist[3] == "")
                    {
                        excelApp.Cells[2, 15] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 15] = "0";
                    }

                    excelApp.Cells[2, 16] = answer_timelist[3];
                    excelApp.Cells[2, 17] = againtimes_list[3];
                    if (answerlist[4] == "")
                    {
                        excelApp.Cells[2, 18] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 18] = "0";
                    }
                    excelApp.Cells[2, 19] = answer_timelist[4];
                    excelApp.Cells[2, 20] = againtimes_list[4];
                    if (answerlist[5] == "")
                    {
                        excelApp.Cells[2, 21] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 21] = "0";
                    }
                    excelApp.Cells[2, 22] = answer_timelist[5];
                    excelApp.Cells[2, 23] = againtimes_list[5];
                    if (answerlist[6] == "")
                    {
                        excelApp.Cells[2, 24] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 24] = "0";
                    }
                    excelApp.Cells[2, 25] = answer_timelist[6];
                    excelApp.Cells[2, 26] = againtimes_list[6];
                    if (answerlist[7] == "")
                    {
                        excelApp.Cells[2, 27] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 27] = "0";
                    }
                    excelApp.Cells[2, 28] = answer_timelist[7];
                    excelApp.Cells[2, 29] = againtimes_list[7];
                    if (answerlist[8] == "")
                    {
                        excelApp.Cells[2, 30] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 30] = "0";
                    }
                    excelApp.Cells[2, 31] = answer_timelist[8];
                    excelApp.Cells[2, 32] = againtimes_list[8];
                    if (answerlist[9] == "")
                    {
                        excelApp.Cells[2, 33] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 33] = "0";
                    }
                    excelApp.Cells[2, 34] = answer_timelist[9];
                    excelApp.Cells[2, 35] = againtimes_list[9];
                    if (answerlist[10] == "")
                    {
                        excelApp.Cells[2, 36] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 36] = "0";
                    }
                    excelApp.Cells[2, 37] = answer_timelist[10];
                    excelApp.Cells[2, 38] = againtimes_list[10];
                    if (answerlist[11] == "")
                    {
                        excelApp.Cells[2, 39] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 39] = "0";
                    }
                    excelApp.Cells[2, 40] = answer_timelist[11];
                    excelApp.Cells[2, 41] = againtimes_list[11];
                    if (answerlist[12] == "")
                    {
                        excelApp.Cells[2, 42] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 42] = "0";
                    }
                    excelApp.Cells[2, 43] = answer_timelist[12];
                    excelApp.Cells[2, 44] = againtimes_list[12];
                    if (answerlist[13] == "")
                    {
                        excelApp.Cells[2, 45] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 45] = "0";
                    }
                    excelApp.Cells[2, 46] = answer_timelist[13];
                    excelApp.Cells[2, 47] = againtimes_list[13];
                    if (answerlist[14] == "")
                    {
                        excelApp.Cells[2, 48] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 48] = "0";
                    }
                    excelApp.Cells[2, 49] = answer_timelist[14];
                    excelApp.Cells[2, 50] = againtimes_list[14];
                    if (answerlist[15] == "")
                    {
                        excelApp.Cells[2, 51] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 51] = "0";
                    }
                    excelApp.Cells[2, 52] = answer_timelist[15];
                    excelApp.Cells[2, 53] = againtimes_list[15];
                    if (answerlist[16] == "")
                    {
                        excelApp.Cells[2, 54] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 54] = "0";
                    }
                    excelApp.Cells[2, 55] = answer_timelist[16];
                    excelApp.Cells[2, 56] = againtimes_list[16];
                    if (answerlist[17] == "")
                    {
                        excelApp.Cells[2, 57] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 57] = "0";
                    }
                    excelApp.Cells[2, 58] = answer_timelist[17];
                    excelApp.Cells[2, 59] = againtimes_list[17];
                    if (answerlist[18] == "3")
                    {
                        excelApp.Cells[2, 60] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 60] = "0";
                    }
                    excelApp.Cells[2, 61] = answer_timelist[18];
                    excelApp.Cells[2, 62] = againtimes_list[18];
                    if (answerlist[19] == "")
                    {
                        excelApp.Cells[2, 63] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 63] = "0";
                    }
                    excelApp.Cells[2, 64] = answer_timelist[19];
                    excelApp.Cells[2, 65] = againtimes_list[19];
                    if (answerlist[20] == "")
                    {
                        excelApp.Cells[2, 66] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 66] = "0";
                    }
                    excelApp.Cells[2, 67] = answer_timelist[20];
                    excelApp.Cells[2, 68] = againtimes_list[20];
                    if (answerlist[21] == "")
                    {
                        excelApp.Cells[2, 69] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 69] = "0";
                    }
                    excelApp.Cells[2, 70] = answer_timelist[21];
                    excelApp.Cells[2, 71] = againtimes_list[21];
                    if (answerlist[22] == "")
                    {
                        excelApp.Cells[2, 72] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 72] = "0";
                    }
                    excelApp.Cells[2, 73] = answer_timelist[22];
                    excelApp.Cells[2, 74] = againtimes_list[22];
                    if (answerlist[23] == "")
                    {
                        excelApp.Cells[2, 75] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 75] = "0";
                    }
                    excelApp.Cells[2, 76] = answer_timelist[23];
                    excelApp.Cells[2, 77] = againtimes_list[23];
                    if (answerlist[24] == "")
                    {
                        excelApp.Cells[2, 78] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 78] = "0";
                    }
                    excelApp.Cells[2, 79] = answer_timelist[24];
                    excelApp.Cells[2, 80] = againtimes_list[24];
                    if (answerlist[25] == "")
                    {
                        excelApp.Cells[2, 81] = "1";
                    }
                    else
                    {
                        excelApp.Cells[2, 81] = "0";
                    }
                    excelApp.Cells[2, 82] = answer_timelist[25];
                    excelApp.Cells[2, 83] = againtimes_list[25];


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


                    //另存活頁簿
                    wBook.SaveAs(pathFile);
                    //Console.WriteLine("儲存文件於 " + Environment.NewLine + pathFile);
                    excelApp.Quit();

                    //釋放Excel資源
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    wBook = null;
                    wSheet = null;
                    //wRange = null;
                    excelApp = null;
                    GC.Collect();
                    Application.Exit();
                    //Console.Read();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("產生表時出錯！" + Environment.NewLine + ex.Message);
                }
            }




            //關閉活頁簿


            //關閉Excel
            
        }
        enum MessageType
        {
            entranceButton,

            Button1,
            Button2,
            Button3,
            nextbutton
        }
    }
}
