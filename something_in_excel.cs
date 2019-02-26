//現在のパスを引き出せる
System.IO.Path.GetDirectoryName(Application.ExecutablePath);


//Rangeでサーチする
Excel.Range colRange = Worksheets.Columns["A:Z"]
Excel.Range resultRange = colRange.Find(
  What: searchString,
  LookIn: Excel.XlFindLookIn.xlValues,
  LookAt: Excel.XlLookAt.xlPart,
  SearchOrder: Excel.XlSearchOrder.xlByRows,
  SearchDirection: Excel.XlSearchDirection.xlNext
);
if (resultRange is null){
  //
}
else
{
  //then handle
}
xlWorkBook.CLose(false);
xlApp.Quit();
releaseObject(xlWorkSheet);
releaseObject(xlWorkBook);
releaseObject(xlApp);

//Excel対象番号をサーチする
var keyword = number;
var hitcell = Excel.Worksheets("Sheet1").Cells.Find(number);
var message = "";
if (hitRange == null){
  //insert new object
}
else
{
   var r = hitcell.Row;
   var c = hitcell.Column;
}


//Excelを閉じる時の最終処理


if (null != this.Excel){
  this.mExcel.DisplayAlerts = false;

  this mExcel.Quit();

  System.Runtime.InteropServices.Marshal.ReleaseComObject(this.mWorkbook);
  System.Runtime.InteropServices.Marshal.ReleaseComObject(this.mExcel);
}
