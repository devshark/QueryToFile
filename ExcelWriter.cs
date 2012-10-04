using System;
using CarlosAg.ExcelXmlWriter;

namespace LoadsReportGen
{
    class ExcelWriter
    {
        public ExcelWriter()
        {
            int ticks = Environment.TickCount;

            // Create the workbook
            Workbook book = new Workbook();
            // Set the author
            book.Properties.Author = "Anthony Lim";

            // Add some style
            WorksheetStyle style = book.Styles.Add("style1");
            style.Font.Bold = true;

            Worksheet sheet = book.Worksheets.Add("SampleSheet");

            WorksheetRow Row0 = sheet.Table.Rows.Add();
            // Add a cell
            Row0.Cells.Add("Hello World", DataType.String, "style1");

            // Save it
            book.Save(@"./test.xls");

            Console.WriteLine("Time:{0}", Environment.TickCount - ticks);
        }
    }
}
