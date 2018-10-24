using Syncfusion.WinForms.Controls;
using Syncfusion.WinForms.DataGrid.Styles;
using System.Windows.Forms;
using Syncfusion.WinForms.DataGrid.Enums;
using Syncfusion.Data;
using System.IO;
using Syncfusion.Pdf.Graphics;
using System.Drawing;
using Syncfusion.Pdf;
using Syncfusion.WinForms.DataGridConverter;
using Syncfusion.WinForms.DataGrid;
using Syncfusion.Pdf.Grid;
using Syncfusion.XlsIO;
using System.Linq;

namespace GettingStarted
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            var data = new OrderInfoCollection();
            sfDataGrid.DataSource = data.OrdersListDetails;
        }

        private void button2_Click(object sender, System.EventArgs e)
        {
            ExcelExportingOptions options = new ExcelExportingOptions();
            
            var excelEngine = this.sfDataGrid.ExportToExcel(sfDataGrid.View, options);


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
    }
}

