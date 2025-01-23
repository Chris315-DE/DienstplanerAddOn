using DienstplanerAddOn.Lib.HelperClass;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DienstplanerAddOn.Lib.Setup
{
    public class MitarbeiterListe : WorksheetBearbeitung
    {
        Worksheet ws;

        private int Buchstabe = 1;
        private int Zahl = 1;
        internal  ExelTabelle tabelle;

        public MitarbeiterListe(Worksheet worksheet) : base(worksheet)
        {
            SetDic(ref KeyCellsManager.MitarbeiterKeyCells);
            tabelle = new ExelTabelle("Mitarbeiter", new List<string> { "MitarbeiterTyp", "Benötigt pro Schicht", "Benötig pro Tag", "Max Schichten am Stück", "Frei nach Block", "Max Schichten pro Monat", "Min Schichten pro Monat" }, new int[] { 3, 2, 2, 2, 2, 2, 2 }, new int[]{1,1,1,1,1,1,1,1 });
            ws = worksheet;
            CreateHeader();
            CreateList();
            CreateFillerText();


        }


        public void CreateHeader()
        {
            var EditCell = ws.Range[ConvertToExcelCell(Buchstabe, Zahl, Buchstabe + 14, 2)];
            EditCell.Merge();
            EditCell.Value = "Setup MitarbeiterListe";
            EditCell.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            EditCell.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
            EditCell.Cells.Font.Bold = true;
            EditCell = ws.Range[ConvertToExcelCell(Buchstabe, Zahl + 2, Buchstabe + 14, Zahl + 4)];
            EditCell.Merge();
            EditCell.Value = "Bitte geben sie hier die verschiedenen MitarbeiterTypen die für den Dienstplan Benötigt werden";
            EditCell.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            EditCell.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
            EditCell.Cells.Font.Bold = true;
            Zahl = 6;

        }


        public void CreateList()
        {
            Zahl = CreateTabelle(ws, Buchstabe, Zahl,
                 Color.LightGray, KeyCellsManager.MitarbeiterKeyCells, tabelle);
        }


        public void CreateFillerText()
        {
            Range edit;
            #region Mitarbeiter
            edit = ws.Range[tabelle["MitarbeiterTyp:0"].Adress];
            edit.Value = "EKS";
            edit = ws.Range[tabelle["MitarbeiterTyp:1"].Adress];
            edit.Value = "Hundeführer";
            edit = ws.Range[tabelle["MitarbeiterTyp:2"].Adress];
            edit.Value = "Consoler";
            #endregion

            #region Pro Schicht
            edit = ws.Range[tabelle["Benötigt pro Schicht:0"].Adress];
            edit.Value = 1;
            edit = ws.Range[tabelle["Benötigt pro Schicht:1"].Adress];
            edit.Value = 2;
            edit = ws.Range[tabelle["Benötigt pro Schicht:2"].Adress];
            edit.Value = 1;
            #endregion

            #region Pro Tag
            edit = ws.Range[tabelle["Benötig pro Tag:0"].Adress];
            edit.Value = 2;
            edit = ws.Range[tabelle["Benötig pro Tag:1"].Adress];
            edit.Value = 4;
            edit = ws.Range[tabelle["Benötig pro Tag:2"].Adress];
            edit.Value = 2;
            #endregion

            #region Max Schichten am Stück

            edit = ws.Range[tabelle["Max Schichten am Stück:0"].Adress];
            edit.Value = 4;
            edit = ws.Range[tabelle["Max Schichten am Stück:1"].Adress];
            edit.Value = 4;
            edit = ws.Range[tabelle["Max Schichten am Stück:2"].Adress];
            edit.Value = 4;

            #endregion

            #region Frei nach Block
            edit = ws.Range[tabelle["Frei nach Block:0"].Adress];
            edit.Value = 3;
            edit = ws.Range[tabelle["Frei nach Block:1"].Adress];
            edit.Value = 3;
            edit = ws.Range[tabelle["Frei nach Block:2"].Adress];
            edit.Value= 3;

            #endregion

            #region Max Schichten pro Monat
            edit = ws.Range[tabelle["Max Schichten pro Monat:0"].Adress];
            edit.Value= 17;
            edit = ws.Range[tabelle["Max Schichten pro Monat:1"].Adress];
            edit.Value= 17;
            edit = ws.Range[tabelle["Max Schichten pro Monat:2"].Adress];
            edit.Value= 17;


            #endregion

            #region Min Schichten pro Monat
            edit = ws.Range[tabelle["Min Schichten pro Monat:0"].Adress];
            edit.Value = 15;
            edit = ws.Range[tabelle["Min Schichten pro Monat:1"].Adress];
            edit.Value = 15;
            edit = ws.Range[tabelle["Min Schichten pro Monat:2"].Adress];
            edit.Value = 15;


            #endregion

        }





        #region Internal

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

        #endregion
    }
}
