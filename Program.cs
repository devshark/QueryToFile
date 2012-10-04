using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.IO;

namespace LoadsReportGen
{
    class Program
    {
        private SqlConnection con = new SqlConnection();
        private Config c;

        static void Main(String[] args)
        {
            main(args);
        }

        static void main(string[] args,bool clear=true)
        {
            //Console.WriteLine(Int32.MaxValue);
            //Console.WriteLine(Int32.Parse("1000.000", System.Globalization.NumberStyles.AllowDecimalPoint));
            //Console.WriteLine(Int32.Parse("1000.000", System.Globalization.NumberStyles.AllowDecimalPoint,(new LoadsXLSCreator.ValueNumberFormat())));

            if(args != null)
                if (args.Count() > 0)
                    foreach(string arg in args)
                        DoSomething(arg,clear);

            Console.WriteLine("Loads Report. What to do?");
            Console.WriteLine("1. Generate Excel Loads File");
            Console.WriteLine("2. Generate XML Loads File");
            Console.WriteLine("3. Generate XLS from Table or View");
            Console.WriteLine("4. Generate XLS from SQL Query Result");
            Console.Write("Action > ");
            String input = Console.ReadLine();
            DoSomething(input);
            //new ExcelWriter();
            //(new Program()).createXMLFile();
        }

        static void DoSomething(string i, bool clear=true)
        {
            if(clear) Console.Clear();
            if(!String.IsNullOrEmpty(i))
            switch (i.Trim())
            {
                case "1":
                    (new Program()).createXLSFile();
                    Console.WriteLine("Done.");
                    main(null);
                    break;
                case "2":
                    (new Program()).createXMLFile();
                    Console.WriteLine("Done.");
                    main(null);
                    break;
                case "3":
                    try
                    {
                        (new Program()).startPrompt();
                        main(null,true);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        main(null,false);
                    }
                    break;
                case "4":
                    try
                    {
                        (new Program()).CreateXLSFromSQL();
                        main(null);
                    }
                    catch (SqlException ex)
                    {
                        Console.WriteLine(ex.Message);
                        main(null, false);
                    }
                    break;
                default:
                    Console.WriteLine("Nothing to do.");
                    main(null);
                    break;
            }
            main(null);
        }

        public Program()
        {
            try{
            this.c = new Config();
            }catch(FileNotFoundException f){
                Console.WriteLine(f.Message);
                Console.ReadLine();
            }
            if(c.Database.IS_BUILTIN)
                this.con = new SqlConnection(String.Format("database={0};server={1};integrated security={2}",this.c.Database.DB,this.c.Database.HOST,this.c.Database.IntegratedSecurity));
            else
                this.con = new SqlConnection(String.Format("database={0};server={1};UID={2};pwd={3};", this.c.Database.DB, this.c.Database.HOST, this.c.Database.UID,this.c.Database.PWD));
            
            this.con.Open();
        }

        void startPrompt()
        {
            Console.Write("Please enter the name of table or view as source."+Environment.NewLine+"Can be comma separated"+Environment.NewLine+" > ");
            string source = Console.ReadLine();
            if (String.IsNullOrEmpty(source.Trim())) return;
            RawTableDump dump = new RawTableDump();
            dump.FetchTables(this.con, source);
        }

        void CreateXLSFromSQL()
        {
            Console.WriteLine("Please enter the SQL Query then press enter.");
            Console.Write(" > ");
            string sql = Console.ReadLine();
            if(String.IsNullOrEmpty(sql.Trim())) return;
            try
            {
                (new RawTableDump()).FetchResult(con, sql);
            }
            catch(SqlException e)
            {
                throw e;
            }
        }

        void createXLSFile()
        {
            using (this.con)
            {
                if(con.State==ConnectionState.Closed) con.Open();
                string sqlAccounts = @"select distinct ACCNO, ACCNAME from View_Kite_LoadReport_ACCNO";
                SqlDataAdapter adp = new SqlDataAdapter(sqlAccounts, con);
                DataSet dst = new DataSet();
                adp.Fill(dst, "ACCOUNTS");

                string sqlAll = "select * from View_Kite_LoadReport_ACCNO order by ACCNO,ACCNAME,CORTEXDATE,CRDPRODUCT,PROGRAMID,ACCTYPE";
                adp = new SqlDataAdapter(sqlAll, con);
                adp.Fill(dst, "ALL");

                LoadsXLSCreator loads = new LoadsXLSCreator(dst.Tables["ACCOUNTS"]);
                loads.CreateFlatSheet(dst.Tables["ALL"]);
                LoadsXLSCreator.CellFieldProperty ACCNO = new LoadsXLSCreator.CellFieldProperty("ACCNO");// "ACCNO", "ACCNAME"\
                //"CORTEXDATE", "CRDPRODUCT", "PROGRAMID", "ACCTYPE", "DESCRIPTION", "AMTBILL"
                loads.HeaderFields((new LoadsXLSCreator.CellFieldProperty("ACCNO")), (new LoadsXLSCreator.CellFieldProperty("ACCNAME")))
                    .DetailFields((new LoadsXLSCreator.CellFieldProperty("CORTEXDATE",CarlosAg.ExcelXmlWriter.DataType.String,"{0:MM/dd/yyyy}")),
                    (new LoadsXLSCreator.CellFieldProperty("CRDPRODUCT")), (new LoadsXLSCreator.CellFieldProperty("PROGRAMID")),
                    (new LoadsXLSCreator.CellFieldProperty("ACCTYPE")),(new LoadsXLSCreator.CellFieldProperty("DESCRIPTION")),
                    (new LoadsXLSCreator.CellFieldProperty("AMTBILL",CarlosAg.ExcelXmlWriter.DataType.Number)))
                    .Create();
            }
        }

        void createXMLFile()
        {
            using(this.con)
            using (StreamWriter sw = System.IO.File.CreateText(string.Format("./Loads{0:ddMMyy}.xml", DateTime.Now)))
            {
                if (con.State==ConnectionState.Closed) con.Open();

                string sqlAccounts = @"select distinct ACCNO, ACCNAME from View_Kite_LoadReport_ACCNO";
                SqlDataAdapter adp = new SqlDataAdapter(sqlAccounts, con);
                DataSet dst = new DataSet();
                adp.Fill(dst, "ACCOUNTS");

                LoadsXMLCreator loads = new LoadsXMLCreator(dst.Tables["ACCOUNTS"]);
                string xml = loads.HeaderFields("ACCNO", "ACCNAME").DetailFields("CORTEXDATE", "CRDPRODUCT", "PROGRAMID", "ACCTYPE", "DESCRIPTION", "AMTBILL").Create();
                sw.Write(xml);
                sw.Close();
            }

        }

        public static DataRowCollection getLoadsDetails(string ACCNO)
        {
            Config c = new Config();
            SqlConnection con;

            if(c.Database.IS_BUILTIN)
                con = new SqlConnection(String.Format("database={0};server={1};integrated security={2}",c.Database.DB,c.Database.HOST,c.Database.IntegratedSecurity));
            else
                con = new SqlConnection(String.Format("database={0};server={1};UID={2};pwd={3};", c.Database.DB, c.Database.HOST, c.Database.UID,c.Database.PWD));

            using (con)
            {
                con.Open();
                string sql = string.Format(@"select CORTEXDATE,CRDPRODUCT,PROGRAMID,ACCTYPE,DESCRIPTION,AMTBILL
from View_Kite_LoadReport_ACCNO
where ACCNO='{0}'
order by CORTEXDATE,CRDPRODUCT,PROGRAMID,ACCTYPE", ACCNO);

                //Console.WriteLine(sql);
                //Console.ReadLine();

                SqlDataAdapter adp = new SqlDataAdapter(sql, con);
                DataSet dst = new DataSet();
                adp.Fill(dst, "LOADS");
                return dst.Tables["LOADS"].Rows;
            }
        }
    }
}
