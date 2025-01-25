using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace DienstplanerAddOn.Lib
{
    public class KeyCellsManager
    {
        public static Dictionary<string, Range> KeyCells;

        public static Dictionary<string, Range> MitarbeiterKeyCells;
        public static Dictionary<string, Range> DiestplanKeyCells;

        public KeyCellsManager()
        {
            KeyCells = new Dictionary<string, Range>();
            MitarbeiterKeyCells = new Dictionary<string, Range>();
            DiestplanKeyCells = new Dictionary<string, Range>();
        }
    }
}
