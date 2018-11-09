 var xl = new ActiveXObject("Excel.Application");
 xl.Visible = true;
 var wb = xl.Workbooks.Open("C:\\Users\\virtualbox\\SASUniversityEdition\\myfolder\\temp\\temp.xlsx");
 var sheet = wb.ActiveSheet;

 wb.Worksheets.Add(after=wb.Sheets(wb.Sheets.Count));
 var pws = wb.ActiveSheet;
 pws.Name="Profit Analysis";
 var pnb=xl.ActiveWindow;
pws.Range("A1").Value ="Pivot Analysis for XXX";
 var pt=wb.PivotCaches().Create(1,wb.sheets("shoe_report").Range("A1").CurrentRegion);
 var pvtTable=pws.PivotTables.Add(PivotCache=pt,TableDestination=pws.Range("A3"),TableName="Profit Analysis");
 var ptformat=pws.PivotTables("Profit Analysis");
 pvtTable.PivotFields("REGION").Orientation = 1;
 pvtTable.PivotFields("PRODUCT").Orientation = 2;
 pvtTable.PivotFields("SALES").Orientation = 4;
 pvtTable.PivotFields("SUM OF SALES").Numberformat="$#,###";
 if (xl.Version > 11) {
 ptformat.TableStyle2 ="PivotStyleLight1";
}

 pws.Columns.Autofit;
 pws.Rows.Autofit;
xl.DisplayAlerts =0;
xl.DisplayAlerts = 1
xl.CutCopyMode = 0;
xl.EnableEvents = 0;
xl = null;
