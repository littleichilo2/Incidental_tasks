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



namespace Inference
{
    public partial class Form1 : Form
    {
        List<string> namelist = new List<string>();
        List<string> birthlist = new List<string>();
        List<string> testlist = new List<string>();
        List<string> answerList = new List<string>();
        string number;
        string gender;
        List<int> titlelist = new List<int>(){
            1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,
            30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54
        };
        List<int> flowerlist = new List<int>(){
            1,6,12,18,23,27,34,41,48
        };
        List<int> leaflist = new List<int>(){
            1,6,12,18,23,27,34,41,48
        };
        List<int> circlelist = new List<int>(){
            3,4,5,8,9,10,14,15,16,17,20,21,22,25,26,29,30,31,
            32,36,37,38,39,43,44,45,46,50,51,52,53
        };
        List<int> Xlist = new List<int>(){
            3,4,5,8,9,10,14,15,16,17,20,21,22,25,26,29,30,31,
            32,36,37,38,39,43,44,45,46,50,51,52,53
        };
        List<int> list_1 = new List<int>() { 11,33,40,47,54 };
        List<int> list_2 = new List<int>() { 11, 33, 40, 47, 54 };
        List<int> list_3 = new List<int>() { 11, 33, 40, 47, 54 };
        List<int> again_buttonlist = new List<int>(){
            3,4,5,8,9,10,11,14,15,16,17,20,21,22,25,26,29,30,31,
            32,33,36,37,38,39,40,43,44,45,46,47,50,51,52,53,54
        };
        List<int> next_buttonlist = new List<int>(){
            2,7,13,19,24,28,35,42,49
        };
        List<string> musicfile_list = new List<string>(){
            "open-i.mp3","practice1.mp3","pi1bgquestion.mp3","question1choice1.mp3",
            "attention-i.mp3","story.mp3","pi-text1.mp3","question2choice2.mp3",
            "pquestion-i11.mp3",
            "pquestion-i12.mp3",
            "pquestion-i13.mp3",
            "practice 2.mp3","pi2bgquestion.mp3",
            "pi-text2.mp3",
            "pquestion-i21.mp3",
            "pquestion-i22.mp3",
            "pquestion-i23.mp3",
            "pquestion-i24.mp3",
            "Istart1.mp3","ibgquestion1.mp3","question1choice1.mp3",
            "attention-i.mp3","story.mp3","i-text1.mp3","question2choice2.mp3",
            "question-i11.mp3",
            "question-i12.mp3",
            "question-i13.mp3",
            "question-i14.mp3",
            "istart2.mp3","ibgquestion2.mp3","question1choice1.mp3",
            "attention-i.mp3","story.mp3","i-text2.mp3","question2choice2.mp3",
            "question-i21.mp3",
            "question-i22.mp3",
            "question-i23.mp3",
            "Istart3.mp3","Ibgquestion3.mp3","question1choice1.mp3",
            "attention-i.mp3","story.mp3","i-text3.mp3","question2choice2.mp3",
            "question-i31.mp3",
            "question-i32.mp3",
            "Istart4.mp3","ibgquesiton4.mp3","question1choice1.mp3",
            "attention-i.mp3","story.mp3","i-text4.mp3","question2choice2.mp3",
            "question-i41.mp3",
            "question-i42.mp3",
            "question-i43.mp3",
            "question-i44.mp3",
            "question-i45.mp3",
            "Istart5.mp3","Ibgquestion5.mp3","question1choice1.mp3",
            "attention-i.mp3","story.mp3","i-text5.mp3","question2choice2.mp3",
            "question-i51.mp3",
            "question-i52.mp3",
            "question-i53.mp3",
            "question-i54.mp3",
            "question-i55.mp3",
            "Istart6.mp3","Ibgquestion6.mp3","question1choice1.mp3",
            "attention-i.mp3","story.mp3","i-text6.mp3","question2choice2.mp3",
            "question-i61.mp3",
            "question-i62.mp3",
            "question-i63.mp3",
            "question-i64.mp3",
            "question-i65.mp3",
            "Istart7.mp3","Ibgquestion7.mp3","question1choice1.mp3",
            "attention-i.mp3","story.mp3","i-text7.mp3","question2choice2.mp3",
            "question-i71.mp3",
            "question-i72.mp3",
            "question-i73.mp3",
            "question-i74.mp3",
            "question-i75.mp3",
            "question1choice1.mp3",
            "question2choice2.mp3"
        };
        List<List<int>> music_list = new List<List<int>>() ;
        

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
            music_list.Add(new List<int> { 0 });
            music_list.Add(new List<int> { 0, 1, 2, 3 });
            music_list.Add(new List<int> { 4, 5, 6, 7 });
            music_list.Add(new List<int> { 8 });
            music_list.Add(new List<int> { 9 });
            music_list.Add(new List<int> { 10 });
            music_list.Add(new List<int> { 11, 12 ,96 });
            music_list.Add(new List<int> { 13,97  });
            music_list.Add(new List<int> { 14  });
            music_list.Add(new List<int> { 15 });
            music_list.Add(new List<int> { 16 });
            music_list.Add(new List<int> { 17 });
            music_list.Add(new List<int> { 18, 19, 20 });
            music_list.Add(new List<int> { 21, 22, 23, 24 });
            music_list.Add(new List<int> { 25 });
            music_list.Add(new List<int> { 26 });
            music_list.Add(new List<int> { 27 });
            music_list.Add(new List<int> { 28 });
            music_list.Add(new List<int> { 29, 30, 31 });
            music_list.Add(new List<int> { 32, 33, 34, 35 });
            music_list.Add(new List<int> { 36 });
            music_list.Add(new List<int> { 37 });
            music_list.Add(new List<int> { 38 });
            music_list.Add(new List<int> { 39, 40, 41 });
            music_list.Add(new List<int> { 42, 43, 44, 45 });
            music_list.Add(new List<int> { 46 });
            music_list.Add(new List<int> { 47 });
            music_list.Add(new List<int> { 48, 49, 50 });
            music_list.Add(new List<int> { 51, 52, 53, 54 });
            music_list.Add(new List<int> { 55 });
            music_list.Add(new List<int> { 56 });
            music_list.Add(new List<int> { 57 });
            music_list.Add(new List<int> { 58 });
            music_list.Add(new List<int> { 59 });
            music_list.Add(new List<int> { 60, 61, 62 });
            music_list.Add(new List<int> { 63, 64, 65, 66 });
            music_list.Add(new List<int> { 67 });
            music_list.Add(new List<int> { 68 });
            music_list.Add(new List<int> { 69 });
            music_list.Add(new List<int> { 70 });
            music_list.Add(new List<int> { 71 });
            music_list.Add(new List<int> { 72, 73, 74 });
            music_list.Add(new List<int> { 75, 76, 77, 78 });
            music_list.Add(new List<int> { 79 });
            music_list.Add(new List<int> { 80 });
            music_list.Add(new List<int> { 81 });
            music_list.Add(new List<int> { 82 });
            music_list.Add(new List<int> { 83 });
            music_list.Add(new List<int> { 84, 85, 86 });
            music_list.Add(new List<int> { 87, 88, 89,90  });
            music_list.Add(new List<int> { 91 });
            music_list.Add(new List<int> { 92 });
            music_list.Add(new List<int> { 93 });
            music_list.Add(new List<int> { 94 });
            music_list.Add(new List<int> { 95 });
        }

        private void button_Click(object sender, EventArgs e, MessageType type)
        {
            
            //add the againtimes into list before the next frame
            if (again_buttonlist.IndexOf(frame) != -1)
                againtimes_list.Add(again);
            again = 0;
            titlelabel.Visible=false;
            flowerbutton.Visible = false;
            leafbutton.Visible = false;

            circlebutton.Visible = false;
            Xbutton.Visible = false;
            Button1.Visible = false;
            Button2.Visible = false;
            Button3.Visible = false;
            againbutton.Visible = false;
            nextbutton.Visible = false;

            frame = frame + 1;
            Console.WriteLine("frame:"+frame);
            
            //music play
            WMPLib.IWMPPlaylist playlist = musicplayer.playlistCollection.newPlaylist("playlist");
            //axWindowsMediaPlayer1
            TimeSpan totaltime = TimeSpan.FromMilliseconds(0);
            if(frame < 55) { 
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
            if (frame == 55)
            {
                endbutton.Visible = true;
            }

            //Record the response from the button
            if (type.Equals(MessageType.circlebutton))
            {
                datetime_now = DateTime.Now;
                answer_timelist.Add((datetime_now - datetimeflag-totaltime).TotalMilliseconds);
                answerlist.Add("O");
                datetimeflag = DateTime.Now;
            }
            else if (type.Equals(MessageType.Xbutton))
            {
                datetime_now = DateTime.Now;
                answer_timelist.Add((datetime_now - datetimeflag - totaltime).TotalMilliseconds);
                answerlist.Add("X");
                datetimeflag = DateTime.Now;
            }
            else if (type.Equals(MessageType.Button1))
            {
                datetime_now = DateTime.Now;
                answer_timelist.Add((datetime_now - datetimeflag - totaltime).TotalMilliseconds);
                answerlist.Add("1");
                datetimeflag = DateTime.Now;
            }
            else if (type.Equals(MessageType.Button2))
            {
                datetime_now = DateTime.Now;
                answer_timelist.Add((datetime_now - datetimeflag - totaltime).TotalMilliseconds);
                answerlist.Add("2");
                datetimeflag = DateTime.Now;
            }
            else if (type.Equals(MessageType.Button3))
            {
                datetime_now = DateTime.Now;
                answer_timelist.Add((datetime_now - datetimeflag - totaltime).TotalMilliseconds);
                answerlist.Add("3");
                datetimeflag = DateTime.Now;
            }
            else if (type.Equals(MessageType.entranceButton))
            {
                teachingpanel.Visible = true;
                datetimeflag = DateTime.Now;
                number =textBox3.Text;
                namelist.Add(textBox1.Text);
                namelist.Add(textBox2.Text);
                birthlist.Add(dateTimePicker1.Text);
                testlist.Add(dateTimePicker2.Text);
                if (radioButton1.Checked)
                {
                    gender = "男";
                }
                else if(radioButton2.Checked)
                {
                    gender = "女";
                }
            }
            else if (type.Equals(MessageType.nextbutton))
            {
                datetimeflag = DateTime.Now;
            }

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

            string pathFile = nowpath +"\\推論理解"+ namelist[0] + namelist[1];


            Excel.Application excelApp = new Excel.Application();
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
                excelApp.Cells[1, 84] = "第二十七題";
                excelApp.Cells[1, 85] = "作答時間";
                excelApp.Cells[1, 86] = "重複次數";
                excelApp.Cells[1, 87] = "第二十八題";
                excelApp.Cells[1, 88] = "作答時間";
                excelApp.Cells[1, 89] = "重複次數";
                excelApp.Cells[1, 90] = "第二十九題";
                excelApp.Cells[1, 91] = "作答時間";
                excelApp.Cells[1, 92] = "重複次數";


                // 設定第1列顏色
                /*wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 2]];
                wRange.Select();
                wRange.Font.Color = ColorTranslator.ToOle(Color.White);
                wRange.Interior.Color = ColorTranslator.ToOle(Color.DimGray);*/

                // 設定第2列資料
                
                excelApp.Cells[2, 1] = number;//流水號
                excelApp.Cells[2, 2] = namelist[0]+namelist[1];//姓名
                excelApp.Cells[2, 3] = birthlist[0];//出生年月
                excelApp.Cells[2, 4] = testlist[0];//施測年月
                excelApp.Cells[2, 5] = gender;//性別
                if (answerlist[0] == "O") {
                    excelApp.Cells[2, 6] = "1";
                }
                else
                {
                    excelApp.Cells[2, 6] = "0";
                }
                excelApp.Cells[2, 7] = answer_timelist[0];
                excelApp.Cells[2, 8] = againtimes_list[0];
                if (answerlist[1] == "X")
                {
                    excelApp.Cells[2, 9] = "1";
                }
                else
                {
                    excelApp.Cells[2, 9] = "0";
                }
                excelApp.Cells[2, 10] = answer_timelist[1];
                excelApp.Cells[2, 11] = againtimes_list[1];
                if (answerlist[2] == "O")
                {
                    excelApp.Cells[2, 12] = "1";
                }
                else
                {
                    excelApp.Cells[2, 12] = "0";
                }

                excelApp.Cells[2, 13] = answer_timelist[2];
                excelApp.Cells[2, 14] = againtimes_list[2];
                if (answerlist[3] == "O")
                {
                    excelApp.Cells[2, 15] = "1";
                }
                else
                {
                    excelApp.Cells[2, 15] = "0";
                }

                excelApp.Cells[2, 16] = answer_timelist[3];
                excelApp.Cells[2, 17] = againtimes_list[3];
                if (answerlist[4] == "O")
                {
                    excelApp.Cells[2, 18] = "1";
                }
                else
                {
                    excelApp.Cells[2, 18] = "0";
                }
                excelApp.Cells[2, 19] = answer_timelist[4];
                excelApp.Cells[2, 20] = againtimes_list[4];
                if (answerlist[5] == "X")
                {
                    excelApp.Cells[2, 21] = "1";
                }
                else
                {
                    excelApp.Cells[2, 21] = "0";
                }
                excelApp.Cells[2, 22] = answer_timelist[5];
                excelApp.Cells[2, 23] = againtimes_list[5];
                if (answerlist[6] == "X")
                {
                    excelApp.Cells[2, 24] = "1";
                }
                else
                {
                    excelApp.Cells[2, 24] = "0";
                }
                excelApp.Cells[2, 25] = answer_timelist[6];
                excelApp.Cells[2, 26] = againtimes_list[6];
                if (answerlist[7] == "O")
                {
                    excelApp.Cells[2, 27] = "1";
                }
                else
                {
                    excelApp.Cells[2, 27] = "0";
                }
                excelApp.Cells[2, 28] = answer_timelist[7];
                excelApp.Cells[2, 29] = againtimes_list[7];
                if (answerlist[8] == "O")
                {
                    excelApp.Cells[2, 30] = "1";
                }
                else
                {
                    excelApp.Cells[2, 30] = "0";
                }
                excelApp.Cells[2, 31] = answer_timelist[8];
                excelApp.Cells[2, 32] = againtimes_list[8];
                if (answerlist[9] == "O")
                {
                    excelApp.Cells[2, 33] = "1";
                }
                else
                {
                    excelApp.Cells[2, 33] = "0";
                }
                excelApp.Cells[2, 34] = answer_timelist[9];
                excelApp.Cells[2, 35] = againtimes_list[9];
                if (answerlist[10] == "X")
                {
                    excelApp.Cells[2, 36] = "1";
                }
                else
                {
                    excelApp.Cells[2, 36] = "0";
                }
                excelApp.Cells[2, 37] = answer_timelist[10];
                excelApp.Cells[2, 38] = againtimes_list[10];
                if (answerlist[11] == "X")
                {
                    excelApp.Cells[2, 39] = "1";
                }
                else
                {
                    excelApp.Cells[2, 39] = "0";
                }
                excelApp.Cells[2, 40] = answer_timelist[11];
                excelApp.Cells[2, 41] = againtimes_list[11];
                if (answerlist[12] == "O")
                {
                    excelApp.Cells[2, 42] = "1";
                }
                else
                {
                    excelApp.Cells[2, 42] = "0";
                }
                excelApp.Cells[2, 43] = answer_timelist[12];
                excelApp.Cells[2, 44] = againtimes_list[12];
                if (answerlist[13] == "1")
                {
                    excelApp.Cells[2, 45] = "1";
                }
                else
                {
                    excelApp.Cells[2, 45] = "0";
                }
                excelApp.Cells[2, 46] = answer_timelist[13];
                excelApp.Cells[2, 47] = againtimes_list[13];
                if (answerlist[14] == "X")
                {
                    excelApp.Cells[2, 48] = "1";
                }
                else
                {
                    excelApp.Cells[2, 48] = "0";
                }
                excelApp.Cells[2, 49] = answer_timelist[14];
                excelApp.Cells[2, 50] = againtimes_list[14];
                if (answerlist[15] == "O")
                {
                    excelApp.Cells[2, 51] = "1";
                }
                else
                {
                    excelApp.Cells[2, 51] = "0";
                }
                excelApp.Cells[2, 52] = answer_timelist[15];
                excelApp.Cells[2, 53] = againtimes_list[15];
                if (answerlist[16] == "X")
                {
                    excelApp.Cells[2, 54] = "1";
                }
                else
                {
                    excelApp.Cells[2, 54] = "0";
                }
                excelApp.Cells[2, 55] = answer_timelist[16];
                excelApp.Cells[2, 56] = againtimes_list[16];
                if (answerlist[17] == "O")
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
                if (answerlist[19] == "O")
                {
                    excelApp.Cells[2, 63] = "1";
                }
                else
                {
                    excelApp.Cells[2, 63] = "0";
                }
                excelApp.Cells[2, 64] = answer_timelist[19];
                excelApp.Cells[2, 65] = againtimes_list[19];
                if (answerlist[20] == "X")
                {
                    excelApp.Cells[2, 66] = "1";
                }
                else
                {
                    excelApp.Cells[2, 66] = "0";
                }
                excelApp.Cells[2, 67] = answer_timelist[20];
                excelApp.Cells[2, 68] = againtimes_list[20];
                if (answerlist[21] == "O")
                {
                    excelApp.Cells[2, 69] = "1";
                }
                else
                {
                    excelApp.Cells[2, 69] = "0";
                }
                excelApp.Cells[2, 70] = answer_timelist[21];
                excelApp.Cells[2, 71] = againtimes_list[21];
                if (answerlist[22] == "O")
                {
                    excelApp.Cells[2, 72] = "1";
                }
                else
                {
                    excelApp.Cells[2, 72] = "0";
                }
                excelApp.Cells[2, 73] = answer_timelist[22];
                excelApp.Cells[2, 74] = againtimes_list[22];
                if (answerlist[23] == "1")
                {
                    excelApp.Cells[2, 75] = "1";
                }
                else
                {
                    excelApp.Cells[2, 75] = "0";
                }
                excelApp.Cells[2, 76] = answer_timelist[23];
                excelApp.Cells[2, 77] = againtimes_list[23];
                if (answerlist[24] == "O")
                {
                    excelApp.Cells[2, 78] = "1";
                }
                else
                {
                    excelApp.Cells[2, 78] = "0";
                }
                excelApp.Cells[2, 79] = answer_timelist[24];
                excelApp.Cells[2, 80] = againtimes_list[24];
                if (answerlist[25] == "X")
                {
                    excelApp.Cells[2, 81] = "1";
                }
                else
                {
                    excelApp.Cells[2, 81] = "0";
                }
                excelApp.Cells[2, 82] = answer_timelist[25];
                excelApp.Cells[2, 83] = againtimes_list[25];
                if (answerlist[26] == "O")
                {
                    excelApp.Cells[2, 84] = "1";
                }
                else
                {
                    excelApp.Cells[2, 84] = "0";
                }
                excelApp.Cells[2, 85] = answer_timelist[26];
                excelApp.Cells[2, 86] = againtimes_list[26];
                if (answerlist[27] == "O")
                {
                    excelApp.Cells[2, 87] = "1";
                }
                else
                {
                    excelApp.Cells[2, 87] = "0";
                }
                excelApp.Cells[2, 88] = answer_timelist[27];
                excelApp.Cells[2, 89] = againtimes_list[27];
                if (answerlist[28] == "1")
                {
                    excelApp.Cells[2, 90] = "1";
                }
                else
                {
                    excelApp.Cells[2, 90] = "0";
                }
                excelApp.Cells[2, 91] = answer_timelist[28];
                excelApp.Cells[2, 92] = againtimes_list[28];

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
                
            }
            catch (Exception ex)
            {
                Console.WriteLine("產生表時出錯！" + Environment.NewLine + ex.Message);
            }

            //關閉活頁簿
            
            
            //關閉Excel
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
