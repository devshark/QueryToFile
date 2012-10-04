using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Text;
using CarlosAg.ExcelXmlWriter;

namespace LoadsReportGen
{
    class RawTableDump
    {
        private Workbook book;
        string filename = string.Format("Data Dump{0:yyyyMMdd}.xls", DateTime.Now);
        Worksheet sheet;
        WorksheetStyle heading;
        WorksheetStyle normal;
        public RawTableDump()
        {
            book = new Workbook(); 

            heading = book.Styles.Add("heading");
            normal = book.Styles.Add("normal");

            heading.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            heading.Font.Bold = true;

            normal.Alignment.Horizontal = StyleHorizontalAlignment.Automatic;
            normal.Font.Bold = false;
            normal.Alignment.Vertical = StyleVerticalAlignment.Center;
        }

        public void FetchResult(SqlConnection con, String SQL)
        {
            if(con.State==ConnectionState.Closed) con.Open();
            try
            {
                using (SqlDataAdapter adp = new SqlDataAdapter(SQL, con))
                {
                    DataSet dst = new DataSet();
                    adp.Fill(dst, "CustomTable");
                    CreateXLS(dst.Tables["CustomTable"]);
                }
            }
            catch
            {
                throw;
            }
        }

        public void FetchTables(SqlConnection con,params String[] TableName)
        {
            foreach (String table in TableName)
            {
                try
                {
                    FetchTable(con, table);
                }
                catch { throw; }
            }
        }

        public void FetchTable(SqlConnection con, String TableName)
        {
            if (con.State == ConnectionState.Closed) con.Open();
            foreach (string table in TableName.Split(",".ToCharArray()))
            {
                String sql = string.Format("select * from {0}", table);
                using (SqlDataAdapter adp = new SqlDataAdapter(sql, con))
                {
                    try
                    {
                        DataSet dst = new DataSet();
                        adp.Fill(dst, table);
                        CreateXLS(dst.Tables[table]);
                        dst.Dispose();
                    }
                    catch
                    {
                        throw;
                    }
                }
            }
            con.Close();
        }

        public void CreateXLS(DataTable table)
        {
            sheet = book.Worksheets.Add(table.TableName);

            WorksheetRow head = sheet.Table.Rows.Add();
            foreach (DataColumn column in table.Columns)
            {
                head.Cells.Add(column.ColumnName, DataType.String, heading.ID);
            }

            foreach (DataRow row in table.Rows)
            {
                WorksheetRow wRow = sheet.Table.Rows.Add();
                foreach(object item in row.ItemArray)
                {
                    wRow.Cells.Add(item.ToString().Trim(), DataType.String, normal.ID);
                }
            }
            book.Save(filename);
        }
    }
}
