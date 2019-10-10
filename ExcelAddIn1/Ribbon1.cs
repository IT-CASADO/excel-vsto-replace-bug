using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet worksheet = workbook.Worksheets[1];
            Range selection = Globals.ThisAddIn.Application.Selection;


            Range range = worksheet.Range[worksheet.Cells[selection.Row, selection.Column], worksheet.Cells[selection.Row + selection.Rows.Count - 1, selection.Column + selection.Columns.Count - 1]];
            range.Interior.Color = XlRgbColor.rgbAquamarine;    // only to verify that correct cells are in the range
            range.Replace(What: " ", Replacement: "");
        }
    }
}
