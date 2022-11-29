using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    public class MST2
    {
        private string _CODE = null;
        private string _NAME = null;
        private DateTime _KIKAN_FROM = DateTime.Now;
        private DateTime _KIKAN_TO = DateTime.Now;

        public MST2()
        {

        }

        public MST2(string code, string name, DateTime kikanFrom, DateTime kikanTo)
        {
            this._CODE = code;
            this._NAME = name;
            this._KIKAN_FROM = kikanFrom;
            this._KIKAN_TO = kikanTo;
        }

        public string CODE
        {
            get { return this._CODE; }
            set { this._CODE = value; }
        }
        public string NAME
        {
            get { return this._NAME; }
            set { this._NAME = value; }
        }
        public DateTime KIKAN_FROM
        {
            get { return this._KIKAN_FROM; }
            set { this._KIKAN_FROM = value; }
        }
        public DateTime KIKAN_TO
        {
            get { return this._KIKAN_TO; }
            set { this._KIKAN_TO = value; }
        }
    }
}
