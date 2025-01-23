using Microsoft.Office.Core;
using System;
using System.Collections;

namespace DienstplanerAddOn.Lib.HelperClass
{
    public class Cells : IEquatable<Cells>
    {
        public Cells(string adress, string key = "", string value = "")
        {
            Adress = adress;
            Key = key;
            Value = value;
        }
        public string Adress { get; set; }
        public string Key { get; set; }
        public string Value { get; set; }


        public Cells()
        {

        }

        public new bool Equals(object x, object y)
        {
            var cell1 = (Cells)x;
            var cell2 = (Cells)y;

            if (cell1 == null || cell2 == null)
                return false;
            if (cell1.Key != cell2.Key)
                return false;
            return true;
        }

        public int GetHashCode(object obj)
        {
            var cell1 = (Cells)(obj);
            return cell1.Key.GetHashCode();
        }

        public bool Equals(Cells other)
        {
            if (other == null) return false;
            if (this.Key != other.Key) return false;
            return true;
        }
    }

}
