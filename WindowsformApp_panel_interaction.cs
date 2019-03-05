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
