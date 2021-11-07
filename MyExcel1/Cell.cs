using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyExcel1
{
    public class Cell : DataGridViewTextBoxCell
    {
        string _value;
        string name;
        string exp;
        bool wasContained = false;
        List<string> cellReference;
        public Cell()
        {
            name = "A1";
            exp = "";
            _value = "0";
            cellReference = new List<string>();
            wasContained = false;
        }

        public string Val
        {
            get { return _value; }
            set { _value = value; }
        }

        public string Exp
        {
            get { return exp; }
            set { exp = value; }
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        public List<string> CellReference
        {
            get { return cellReference; }
            set { cellReference = value; }
        }

        public bool WasContained
        {
            get { return wasContained; }
            set { wasContained = value; }
        }
    }
}