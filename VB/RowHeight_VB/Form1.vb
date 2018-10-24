Imports Microsoft.VisualBasic
Imports Syncfusion.WinForms.Controls
Imports Syncfusion.WinForms.DataGrid.Styles
Imports System.Windows.Forms
Imports Syncfusion.WinForms.DataGrid.Enums
Imports Syncfusion.Data
Imports System.IO
Imports Syncfusion.Pdf.Graphics
Imports System.Drawing
Imports Syncfusion.Pdf
Imports Syncfusion.WinForms.DataGridConverter
Imports Syncfusion.WinForms.DataGrid
Imports Syncfusion.Pdf.Grid
Imports Syncfusion.XlsIO
Imports System.Linq

Namespace GettingStarted
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
			Dim data = New OrderInfoCollection()
			sfDataGrid.DataSource = data.OrdersListDetails
		End Sub

		Private Sub button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button2.Click
			Dim options As New ExcelExportingOptions()

			Dim excelEngine = Me.sfDataGrid.ExportToExcel(sfDataGrid.View, options)


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
	End Class
End Namespace

