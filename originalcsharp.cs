using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
        int frame;
        public Form1()
        {
            InitializeComponent();

            

        }

        private void entranceButton_Click(object sender, EventArgs e)
        {
            frame = frame + 1;
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
        }
    }
}
