using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
namespace DienstplanerAddOn.Lib
{
    public class ObsoletCreateMitarbeiterListe : WorksheetBearbeitung
    {
        Worksheet ws;

        private int Buchstabe = 1;
        private int Zahl = 1;

        public ObsoletCreateMitarbeiterListe(Worksheet activesheet) : base(activesheet)
        {
            SetDic(ref KeyCellsManager.MitarbeiterKeyCells);
            ws = activesheet;
            CreateHeadder();



          
     


        

        }

        private void CreateHeadder()
        {
            if (ws == null)
                return;

            string test = ConvertToExcelCell(Buchstabe, Zahl, Buchstabe + 13, Zahl + 1);

            ws.Range[ConvertToExcelCell(Buchstabe, Zahl, Buchstabe + 13, Zahl + 1)].Merge();
            Range cell = ws.Range[ConvertToExcelCell(Buchstabe, Zahl)];
            cell.Value = "Mitarbeiter Setup";
            cell.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            cell.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
            Zahl += 2;
            cell = ws.Range[ConvertToExcelCell(Buchstabe, Zahl, Buchstabe + 13, Zahl)];
            cell.Merge();
            cell.Interior.Color = Color.Black;
            Zahl += 2;


        }

        public ObsoletCreateMitarbeiterListe CreateMitarbeiterInfo(string name, string key)
        {


            List<string> Settings = new List<string>() 
            {
                "Bezeichnung",
                "Anzahl Mitarbeiter",
                "Max Schichten am Stück",
                "Max Schichten je Monat",
                "Min Schichten je Monat",
                "Min Tage frei nach Schichtblock",
            };


            Range cell = ws.Range[ConvertToExcelCell(Buchstabe, Zahl, Buchstabe + 3, Zahl)];
            cell.Merge();
            cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            cell.VerticalAlignment = XlVAlign.xlVAlignCenter;
            cell.Value = name;
            ErstelleKeyDicEintrag(key + name, cell);
            Zahl++;
            foreach (var setting in Settings)
            {
                cell = ws.Range[ConvertToExcelCell(Buchstabe, Zahl, Buchstabe+1, Zahl)];
                cell.Merge();
                cell.Value = setting;
                cell.Style.WrapText = true;
                Range keycell = ws.Range[ConvertToExcelCell(Buchstabe+2, Zahl, Buchstabe+3, Zahl)];
                keycell.Merge();
                ErstelleKeyDicEintrag(key + setting, keycell);
                keycell.Interior.Color = Color.LightGray;
                if(setting == "Bezeichnung")
                {
                    keycell.Value = name;
                }
                Zahl++;
                if (setting == Settings[5])
                {
                    cell.RowHeight = cell.RowHeight + cell.RowHeight;
                }
            }

            Range formelr = GetRange(key + Settings[0]);
            Range target = GetRange(key + name);

            ws.Range[ConvertToExcelCell(target.Column,target.Row)].Formula = $"=CONCATENATE({ConvertToExcelCell(formelr.Column, formelr.Row)},\"\")";
            return this;
        }

        internal override void ErstelleKeyDicEintrag(string key, Range value)
        {
            if ((bool)KeyCellsManager.MitarbeiterKeyCells?.ContainsKey(key))
                return;
            KeyCellsManager.MitarbeiterKeyCells.Add(key, value);
        }

        internal override Range GetRange(string key)
        {
            KeyCellsManager.MitarbeiterKeyCells.TryGetValue(key, out var cell);
            return cell;
        }

    }
}
