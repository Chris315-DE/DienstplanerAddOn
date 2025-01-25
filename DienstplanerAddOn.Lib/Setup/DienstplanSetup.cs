using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DienstplanerAddOn.Lib.Setup
{
    public class DienstplanSetup : WorksheetBearbeitung
    {
        public DienstplanSetup(Worksheet worksheet) : base(worksheet)
        {
            SetDic(ref KeyCellsManager.DiestplanKeyCells);
            CreateHeader("Dienstplan Setup", "Work in Progress...");
        }

        internal override void ErstelleKeyDicEintrag(string key, Range value)
        {
            if ((bool)KeyCellsManager.DiestplanKeyCells?.ContainsKey(key))
                return;
            KeyCellsManager.MitarbeiterKeyCells.Add(key, value);
        }

        internal override Range GetRange(string key)
        {
            KeyCellsManager.DiestplanKeyCells.TryGetValue(key, out var cell);
            return cell;
        }
    }
}
