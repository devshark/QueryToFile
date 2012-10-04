using System;
using System.Collections.Generic;
using System.Data;

namespace LoadsReportGen
{
    class LoadsXMLCreator
    {
        private List<string> __headers = new List<string>();
        private List<string> __details = new List<string>();
        private DataTable tblAccounts = new DataTable();
        private const string NODESPLACEHOLDER = "[^NODES$]";

        public LoadsXMLCreator(DataTable AccountsTbl)
        {
            this.tblAccounts = AccountsTbl;
        }

        public LoadsXMLCreator HeaderFields(params string[] hField)
        {
            foreach (string fld in hField) this.__headers.Add(fld);
            return this;
        }

        public LoadsXMLCreator DetailFields(params string[] dFields)
        {
            foreach (string fld in dFields) this.__details.Add(fld);
            return this;
        }

        public string Create()
        {
            string prevACCNO = string.Empty; string XML = string.Empty;

            string details = string.Empty;

            foreach (DataRow row in this.tblAccounts.Rows)
            {
                XML += createHeader(row).Replace(NODESPLACEHOLDER, createDetails(Program.getLoadsDetails(row[0].ToString())));
            }
            return startXML() + parentWrapper(XML);
        }

        protected string startXML()
        {
            return "<?xml version=\"1.0\"?>";
        }

        protected string parentWrapper(string XMLNodes)
        {
            return "<LOADS>" + XMLNodes + "</LOADS>";
        }

        protected string createHeader(DataRow row)
        {
            string ACCOUNT = "<ACCOUNT {0}>" + NODESPLACEHOLDER + "</ACCOUNT>"; string h = string.Empty;
            foreach (string head in this.__headers)
            {
                h += string.Format("{0}=\"{1}\" ", head.Trim(), row[head].ToString().Trim());
            }
            return string.Format(ACCOUNT, h).ToUpper();
        }

        protected string createDetails(DataRowCollection rows)
        {
            string XML = string.Empty;
            foreach (DataRow row in rows)
            {
                XML += "<LOAD>";
                foreach (string detail in this.__details)
                {
                    XML += string.Format("<{0}>{1}</{0}>", detail.Trim(), row[detail].ToString().Trim());
                }
                XML += "</LOAD>";
            }
            return XML.ToUpper();
        }

    }
}
