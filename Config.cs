using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using Anthony.Lim;
using System.IO;

namespace LoadsReportGen
{
    class Config
    {
        protected string configFile = "./config.js";
        private string strConfig;
        protected Hashtable config;
        private DATABASE _db;
        public DATABASE Database { get { return this._db; } }

        public Hashtable CONFIG { get { return this.config; } }

        public Config()
        {
            this.LoadConfig();
            this._db = new DATABASE(this.config);
        }

        public void LoadConfig()
        {
            if (System.IO.File.Exists(this.configFile))
            {
                this.strConfig = System.IO.File.ReadAllText(this.configFile);
                this.config = (Hashtable)JSON.JsonDecode(this.strConfig);
            }
            else
            {
                throw new FileNotFoundException("The config file was not found. Nothing to do.");
            }
        }

        public class DATABASE
        {
            private string _host;
            private string _db;
            private bool _isBuiltin;
            private string _uid;
            private string _pwd;
            private string _security;

            public String HOST { get { return this._host; } }
            public String DB { get { return this._db; } }
            public bool IS_BUILTIN { get { return this._isBuiltin; } }
            public String UID { get { return this._uid; } }
            public String PWD { get { return this._pwd; } }
            public String IntegratedSecurity { get { return this._security; } }

            public DATABASE(Hashtable config)
            {
                this._host = (string)config["host"];
                this._db = (string)config["db"];
                this._isBuiltin = (bool)config["builtin"];
                this._uid = (string)config["uid"];
                this._pwd = (string)config["pwd"];
                this._security = (string)config["security"];
            }
        }
    }
}
