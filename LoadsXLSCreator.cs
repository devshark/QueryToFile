using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using CarlosAg.ExcelXmlWriter;
using System.Globalization;

namespace LoadsReportGen
{
    class LoadsXLSCreator
    {
        private List<CellFieldProperty> __headers = new List<CellFieldProperty>();
        private List<CellFieldProperty> __details = new List<CellFieldProperty>();
        private DataTable tblAccounts = new DataTable();
        Workbook book = new Workbook();
        WorksheetStyle headings = null;
        WorksheetStyle coloredHeading = null;
        WorksheetStyle normal = null;
        WorksheetStyle rightAlign = null;
        private string filename;

        public LoadsXLSCreator(DataTable AccountsTbl)
        {
            this.tblAccounts = AccountsTbl;
            this.filename = string.Format("./loads{0:ddMMyy}.xls", DateTime.Now); 

            book.Properties.Author = "Anthony Lim";
            book.Properties.Company = "Kitesystems";

            headings = book.Styles.Add("Heading1");
            headings.Font.Bold = true;
            headings.Font.Size = 12;
            headings.Alignment.Horizontal = StyleHorizontalAlignment.Center;

            coloredHeading = book.Styles.Add("ColoredHeading");
            coloredHeading.Font.Size = 12;
            coloredHeading.Font.Color = "#FFFFFF";
            coloredHeading.Interior.Color = "#55AAFF";
            coloredHeading.Interior.Pattern = StyleInteriorPattern.Solid;
            coloredHeading.Font.Italic = true;
            coloredHeading.Font.Bold = true;
            coloredHeading.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            
            this.normal = book.Styles.Add("Default");
            this.normal.Font.Size = 10;
            this.normal.Font.Bold = false;

            this.rightAlign = book.Styles.Add("RightAligned");
            this.rightAlign.Font.Size = 10;
            this.normal.Font.Bold = false;
            this.rightAlign.Alignment.Horizontal = StyleHorizontalAlignment.Right;
        }

        public LoadsXLSCreator HeaderFields(params CellFieldProperty[] hFields)
        {
            foreach (CellFieldProperty fld in hFields) this.__headers.Add(fld);
            return this;
        }

        public LoadsXLSCreator DetailFields(params CellFieldProperty[] dFields)
        {
            foreach (CellFieldProperty fld in dFields) this.__details.Add(fld);
            return this;
        }

        public static float getNumericEquivalent(string val)
        {
            float fin; float.TryParse(val, out fin); return fin;
        }

        public void Create()
        {
            Worksheet matrix = book.Worksheets.Add("Matrix");

            foreach (DataRow row in this.tblAccounts.Rows)
            {
                WorksheetRow head = matrix.Table.Rows.Add();
                foreach (CellFieldProperty h in this.__headers)
                {
                    head.Cells.Add(h.FieldName, DataType.String, this.coloredHeading.ID);
                    head.Cells.Add(h.getFormatted(row),h.dataType,normal.ID);
                }

                WorksheetRow detHead = matrix.Table.Rows.Add();
                foreach (CellFieldProperty d in this.__details)
                    detHead.Cells.Add(d.FieldName,DataType.String,headings.ID);

                float totals = 0.00f;

                DataRowCollection details = Program.getLoadsDetails(row[0].ToString());
                foreach (DataRow detail in details)
                {
                    WorksheetRow det = matrix.Table.Rows.Add();
                    foreach (CellFieldProperty d in this.__details)
                    {
                        det.Cells.Add(d.getFormatted(detail), d.dataType, normal.ID);
                        if (d.dataType == DataType.Number || d.dataType == DataType.Integer)
                            totals += getNumericEquivalent(detail[d.FieldName].ToString());
                    }
                }
                WorksheetRow RowTotal = matrix.Table.Rows.Add();
                foreach (CellFieldProperty f in this.__details)
                    if (f.dataType == DataType.Integer || f.dataType == DataType.Number)
                        RowTotal.Cells.Add(totals.ToString("0.####"), f.dataType, headings.ID);
                    else
                        RowTotal.Cells.Add();

                matrix.Table.Rows.Add();
            }

            book.Save(this.filename);
        }

        public void CreateFlatSheet(DataTable dt)
        {
            Worksheet flat = book.Worksheets.Add("Flat");

            WorksheetRow head = flat.Table.Rows.Add();
            foreach (DataColumn col in dt.Columns)
                head.Cells.Add(col.Caption, DataType.String, this.headings.ID);

            foreach (DataRow row in dt.Rows)
            {
                WorksheetRow line = flat.Table.Rows.Add();
                for (int i = 0; i < row.ItemArray.Length; i++)
                {
                    //float val; //int du; ValueNumberFormat v = new ValueNumberFormat();
                    //if (float.TryParse(row[i].ToString(),NumberStyles.Integer,(new ValueNumberFormat()), out val) /*&& Int32.TryParse(row[i].ToString(), NumberStyles.Number,v, out du)*/)
                    if(i==8 || i==6)
                        line.Cells.Add(row[i].ToString(), i==8 ? DataType.Number : DataType.String, i==8 ? normal.ID : rightAlign.ID);
                    else
                        line.Cells.Add(row[i].ToString(),DataType.String,normal.ID);
                }
            }
            book.Save(this.filename);
        }

        public class CellFieldProperty
        {
            public DataType dataType { get; set; }
            public  String FieldName { get; set; }
            public String Format { get; set; }

            public CellFieldProperty(string fieldname, DataType dt = DataType.String, String Format = "{0}")
            {
                this.FieldName = fieldname;
                this.dataType = dt;
                this.Format = Format;
            }

            public string getFormatted(string value)
            {
                return String.Format(this.Format, value);
            }

            public string getFormatted(DataRow row)
            {
                return String.Format(this.Format, row[this.FieldName]);
            }

        }
        public class ValueNumberFormat : IFormatProvider,ICustomFormatter
        {
            private const int ACCT_LENGTH = 6;

            public object GetFormat(Type formatType)
            {
                if (formatType == typeof(ICustomFormatter))
                    return this;
                else
                    return null;
            }

            public string Format(string fmt, object arg, IFormatProvider formatProvider)
            {
                // Provide default formatting if arg is not an Int64.
                if (arg.GetType() != typeof(Int32))
                    try
                    {
                        return HandleOtherFormats(fmt, arg);
                    }
                    catch (FormatException e)
                    {
                        throw new FormatException(String.Format("The format of '{0}' is invalid.", fmt), e);
                    }

                // Provide default formatting for unsupported format strings.
                string ufmt = fmt.ToUpper(CultureInfo.InvariantCulture);
                if (!(ufmt == "H" || ufmt == "I"))
                    try
                    {
                        return HandleOtherFormats(fmt, arg);
                    }
                    catch (FormatException e)
                    {
                        throw new FormatException(String.Format("The format of '{0}' is invalid.", fmt), e);
                    }

                // Convert argument to a string.
                string result = arg.ToString();

                // If account number is less than 12 characters, pad with leading zeroes.
                if (result.Length < ACCT_LENGTH)
                    result = result.PadLeft(ACCT_LENGTH, '0');
                // If account number is more than 12 characters, truncate to 12 characters.
                if (result.Length > ACCT_LENGTH)
                    result = result.Substring(0, ACCT_LENGTH);

                if (ufmt == "I")                    // Integer-only format. 
                    return result;
                // Add hyphens for H format specifier.
                else                                         // Hyphenated format.
                    return result.Substring(0, 5) + "-" + result.Substring(5, 3) + "-" + result.Substring(8);
            }

            private string HandleOtherFormats(string format, object arg)
            {
                if (arg is IFormattable)
                    return ((IFormattable)arg).ToString(format, CultureInfo.CurrentCulture);
                else if (arg != null)
                    return arg.ToString();
                else
                    return String.Empty;
            }
        }
    }
}
