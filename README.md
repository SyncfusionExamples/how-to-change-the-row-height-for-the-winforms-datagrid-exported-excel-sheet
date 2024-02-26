# How to change the row height for the exported excel sheet in WinForms DataGrid (SfDataGrid)

## Change the row height

You can change the row height of the exported excel sheet using <b>Worksheets.UsedRange.RowHeight</b> property.

## C#

```C#
private void ExportToExcel_Click(object sender, System.EventArgs e)
{
    ExcelExportingOptions options = new ExcelExportingOptions();
    var excelEngine = sfDataGrid.ExportToExcel(sfDataGrid.View, options);
    var workBook = excelEngine.Excel.Workbooks[0];
    
    //Set row height.
    workBook.Worksheets[0].UsedRange.RowHeight = 30;
 
    SaveFileDialog sfd = new SaveFileDialog
    {
        FilterIndex = 2,
        Filter = "Excel 97 to 2003 Files(*.xls)|*.xls|Excel 2007 to 2010 Files(*.xlsx)|*.xlsx",
        FileName = "Book1"
    };
 
    if (sfd.ShowDialog() == DialogResult.OK)
    {
        using (Stream stream = sfd.OpenFile())
        {
            if (sfd.FilterIndex == 1)
                workBook.Version = ExcelVersion.Excel97to2003;
            else
                workBook.Version = ExcelVersion.Excel2010;
            workBook.SaveAs(stream);
        }
 
        //Message box confirmation to view the created spreadsheet.
        if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created", MessageBoxButtons.OKCancel) == DialogResult.OK)
        {
            //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
            System.Diagnostics.Process.Start(sfd.FileName);
        }
    }
}
```

## VB

```VB
Private Sub ExportToExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button2.Click
    Dim options As New ExcelExportingOptions()
    Dim excelEngine = sfDataGrid.ExportToExcel(sfDataGrid.View, options)
    Dim workBook = excelEngine.Excel.Workbooks(0)
       
    'Set row height.
    workBook.Worksheets(0).UsedRange.RowHeight = 30
 
    Dim sfd As SaveFileDialog = New SaveFileDialog With {.FilterIndex = 2, .Filter = "Excel 97 to 2003 Files(*.xls)|*.xls|Excel 2007 to 2010 Files(*.xlsx)|*.xlsx", .FileName = "Book1"}
 
    If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
       Using stream As Stream = sfd.OpenFile()
       If sfd.FilterIndex = 1 Then
          workBook.Version = ExcelVersion.Excel97to2003
       Else
          workBook.Version = ExcelVersion.Excel2010
       End If
           workBook.SaveAs(stream)
       End Using
 
       'Message box confirmation to view the created spreadsheet.
       If MessageBox.Show("Do you want to view the workbook?", "Workbook has been created", MessageBoxButtons.OKCancel) = System.Windows.Forms.DialogResult.OK Then
          'Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
          System.Diagnostics.Process.Start(sfd.FileName)
       End If
    End If
End Sub
```

You can use <b>AutofitRows</b> method to adjust the row height of the exported excel sheet based on the content.

## C#

```C#
private void ExportToExcel_Click(object sender, System.EventArgs e)
{
    ExcelExportingOptions options = new ExcelExportingOptions();
    var excelEngine = sfDataGrid.ExportToExcel(sfDataGrid.View, options);
    var workBook = excelEngine.Excel.Workbooks[0];
 
    //Row height will be set based on the content.
    workBook.Worksheets[0].UsedRange.AutofitRows();
 
    SaveFileDialog sfd = new SaveFileDialog
    {
        FilterIndex = 2,
        Filter = "Excel 97 to 2003 Files(*.xls)|*.xls|Excel 2007 to 2010 Files(*.xlsx)|*.xlsx",
        FileName = "Book1"
    };
 
    if (sfd.ShowDialog() == DialogResult.OK)
    {
        using (Stream stream = sfd.OpenFile())
        {
            if (sfd.FilterIndex == 1)
                workBook.Version = ExcelVersion.Excel97to2003;
            else
                workBook.Version = ExcelVersion.Excel2010;
            workBook.SaveAs(stream);
        }
 
        //Message box confirmation to view the created spreadsheet.
        if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created", MessageBoxButtons.OKCancel) == DialogResult.OK)
        {
            //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
            System.Diagnostics.Process.Start(sfd.FileName);
        }
    }
}
```

## VB

```VB
Private Sub ExportToExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button2.Click
    Dim options As New ExcelExportingOptions()
    Dim excelEngine = sfDataGrid.ExportToExcel(sfDataGrid.View, options)
    Dim workBook = excelEngine.Excel.Workbooks(0)
 
    'Row height will be set based on the content.
    workBook.Worksheets(0).UsedRange.AutofitRows()
 
    Dim sfd As SaveFileDialog = New SaveFileDialog With {.FilterIndex = 2, .Filter = "Excel 97 to 2003 Files(*.xls)|*.xls|Excel 2007 to 2010 Files(*.xlsx)|*.xlsx", .FileName = "Book1"}
 
     If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        Using stream As Stream = sfd.OpenFile()
 If sfd.FilterIndex = 1 Then
    workBook.Version = ExcelVersion.Excel97to2003
 Else
    workBook.Version = ExcelVersion.Excel2010
 End If
              workBook.SaveAs(stream)
        End Using
 
        'Message box confirmation to view the created spreadsheet.
        If MessageBox.Show("Do you want to view the workbook?", "Workbook has been created", MessageBoxButtons.OKCancel) = System.Windows.Forms.DialogResult.OK Then
           'Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
           System.Diagnostics.Process.Start(sfd.FileName)
        End If
     End If
End Sub
```