
 var xl = new ActiveXObject("Excel.Application");
 xl.Visible = true;
 var wb = xl.Workbooks.Open("/folders/myshortcuts/myfolder/temp/temp.xlsx");
 var sheet = wb.ActiveSheet;

 wb.Worksheets.Add(after=wb.Sheets(wb.Sheets.Count));
 var pws = wb.ActiveSheet;
 pws.Name="Profit Analysis Chart";
 var pnb=xl.ActiveWindow;
pws.Range("A1").Value ="Pivot Analysis for XXX";
 var pt=wb.PivotCaches().Create(1,wb.sheets("shoe_report").Range("A1").CurrentRegion);
 var pvtTable=pws.PivotTables.Add(PivotCache=pt,TableDestination=pws.Range("A3"),TableName="Profit Analysis Chart");
 var ptformat=pws.PivotTables("Profit Analysis Chart");
 pvtTable.PivotFields("REGION").Orientation = 1;
 pvtTable.PivotFields("PRODUCT").Orientation = 2;
 pvtTable.PivotFields("SALES").Orientation = 4;
 pvtTable.PivotFields("SUM OF SALES").Numberformat="$#,###";
 if (xl.Version > 11) {
 ptformat.TableStyle2 ="PivotStyleLight1";
}

 pws.Columns.Autofit;
 pws.Rows.Autofit;

// Adding new pivot table chart sheet 
 var ch = wb.Charts.Add(wb.Sheets(wb.Sheets.Count));
 var ach=wb.ActiveChart;
 var prang=wb.sheets("Profit Analysis Chart").UsedRange;
 ach.SetSourceData(prang);
 ach.ChartType ="51";
 ach.Name="Profit Analysis Charts";

xl.DisplayAlerts =0;
xl.DisplayAlerts = 1
xl.CutCopyMode = 0;
xl.EnableEvents = 0;
xl = null;

