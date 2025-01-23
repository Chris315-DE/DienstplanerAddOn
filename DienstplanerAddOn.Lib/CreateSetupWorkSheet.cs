using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace DienstplanerAddOn.Lib
{
    public class CreateSetupWorkSheet
    {
        private Dictionary<string, Range> KeyCells;
        Worksheet WS;

        private int startBuchstabe = 1;
        private int startZahl = 1;
        private bool backup = false;
        private int backupZahl;
        private int backupBuchstabe;

        private List<string> Titels = new List<string>()
            {
                "Erster Arbeitstag der Woche:",
                "Letzter Arbeitstag der Woche:",
                "Arbeitstage in der Woche:",
                "Schichten pro Tag:",
                "Mitarbeiter typ A pro Schicht:",
                "Mitarbeiter typ B pro Schicht:",
                "Mitarbeiter typ C pro Schicht:",
            };

        private List<string> Infos = new List<string>()
            {
                "Auswahl: Mo,Di,Mi,Do,Fr,Sa,So",
                "Auswahl: Mo,Di,Mi,Do,Fr,Sa,So",
                "",
                "",
                "Mitarbeiter A pro Tag pro Schicht:",
                "Mitarbeiter B pro Tag pro Schicht:",
                "Mitarbeiter C pro Tag pro Schicht:",
            };


        private List<string> MitarbeiterInfos = new List<string>()
        {
            "Bezeichnung",
            "Anzahl MA",
            "Max Schichten am Stück",
            "Max Schichten pro Monat",
            "Min Schichten pro Monat",
            "Min Tage frei nach Schichtblock",
        };


        public CreateSetupWorkSheet(Worksheet activeWorksheet)
        {
            KeyCells = KeyCellsManager.KeyCells;
            WS = activeWorksheet;

            CreateHeadder();
            //A3 
            CreateStundenInfos(startBuchstabe, startZahl);
            CreateMitarbeiterInfo("Mitarbeiter typ A", Titels[4], 1);
            CreateMitarbeiterInfo("Mitarbeiter typ B", Titels[5], 2);
            CreateMitarbeiterInfo("Mitarbeiter typ C", Titels[6], 3);
        }

        private void CreateHeadder()
        {
            if (WS == null)
                return;
            Excel.Worksheet ws = WS;
            ws.Activate();



            ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}:{ConverttoChar(startBuchstabe + 13)}{startZahl + 1}"].Merge();
            ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"].Value = "Setup";
            ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            startZahl = 3;
        }

        private void CreateStundenInfos(int startrow, int startcol)
        {
            if (WS == null)
                return;
            Worksheet ws = WS;
            ws.Activate();





            int startBuchstabe = startrow;
            int startZahl = startcol;
            //A3:C3

            foreach (var titel in Titels)
            {
                ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}:{ConverttoChar(startBuchstabe + 2)}{startZahl}"].Merge();
                ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"].Value = titel;
                ws.Range[$"{ConverttoChar(startBuchstabe + 3)}{startZahl}"].Interior.Color = Color.LightGray;
                ErstelleDicLinks(titel, ws.Range[$"{ConverttoChar(startBuchstabe + 3)}{startZahl}"]);
                startZahl++;


            }



            //A10
            ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}:{ConverttoChar(startBuchstabe + 2)}{startZahl}"].Merge();
            startZahl++;

            //A11
            ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}:{ConverttoChar(startBuchstabe + 7)}{startZahl}"].Merge();
            ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"].Interior.Color = Color.Black;

            startBuchstabe = startrow + 3;
            startZahl = startcol;



            startBuchstabe++;

            foreach (var info in Infos)
            {
                ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}:{ConverttoChar(startBuchstabe + 2)}{startZahl}"].Merge();
                ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"].Value = info;

                ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                startZahl++;
            }

            //Startzahl = 10
            //Buchstabe 5 = E 

            startZahl -= 3;
            startBuchstabe += 3;
            int a = startBuchstabe;

            Range schichtenProTag = KeyCells[Titels[3]];
            Range MitarbeiterTypA = KeyCells[Titels[4]];
            Range MitarbeiterTypB = KeyCells[Titels[5]];
            Range MitarbeiterTypC = KeyCells[Titels[6]];



            ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"].Formula = $"={ConverttoChar(schichtenProTag.Cells.Column)}{schichtenProTag.Cells.Row}*{ConverttoChar(MitarbeiterTypA.Cells.Column)}{MitarbeiterTypA.Cells.Row}";
            ErstelleDicLinks(Infos[4], ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"]);

            startZahl++;

            ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"].Formula =
                $"={ConverttoChar(schichtenProTag.Cells.Column)}{schichtenProTag.Cells.Row}*{ConverttoChar(MitarbeiterTypB.Cells.Column)}{MitarbeiterTypB.Cells.Row}";
            ErstelleDicLinks(Infos[5], ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"]);
            startZahl++;
            ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"].Formula =
                 $"={ConverttoChar(schichtenProTag.Cells.Column)}{schichtenProTag.Cells.Row}*{ConverttoChar(MitarbeiterTypC.Cells.Column)}{MitarbeiterTypC.Cells.Row}";
            ErstelleDicLinks(Infos[6], ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"]);
            startZahl++;
            ws.Range[$"{ConverttoChar(startBuchstabe - 1)}{startZahl}"].Value = "Gesammt:";


            Range SummeA = KeyCells[Infos[4]];
            Range SummeB = KeyCells[Infos[5]];
            Range SummeC = KeyCells[Infos[6]];

            ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"].Formula =
                $"={ConverttoChar(SummeA.Column)}{SummeA.Row}+{ConverttoChar(SummeB.Column)}{SummeB.Row}+{ConverttoChar(SummeC.Column)}{SummeC.Row}";

            this.startBuchstabe = startBuchstabe;
            this.startZahl = startZahl;


        }

        private void CreateMitarbeiterInfo(string name, string key, int id)
        {
            if (WS == null)
                return;
            Excel.Worksheet ws = WS;
            startBuchstabe = 1;
            startZahl += 2;

            if (!backup)
            {
                backupBuchstabe = startBuchstabe;
                backupZahl = startZahl;
                backup = true;
            }


            //Header
            ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}:{ConverttoChar(startBuchstabe + 3)}{startZahl}"].Merge();
            ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"].Formula = $"=CONCATENATE({ConverttoChar(startBuchstabe + 2)}{startZahl + 1},\" Info ID: {id}\")";

            //ReName
            Range rename = KeyCells[key];
            ws.Range[$"{ConverttoChar(1)}{rename.Row}"].Formula = $"=CONCATENATE({ConverttoChar(startBuchstabe + 2)}{startZahl + 1},\" pro Schicht:\")";


            startZahl++;

            int counter = 0;
            foreach (var maInfo in MitarbeiterInfos)
            {

                if (MitarbeiterInfos.IndexOf(maInfo) == MitarbeiterInfos.Count - 1)
                {

                    ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}:{ConverttoChar(startBuchstabe + 1)}{startZahl + 1}"].Merge();
                    ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"].Value = maInfo;
                    ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"].Style.WrapText = true;
                    ws.Range[$"{ConverttoChar(startBuchstabe + 2)}{startZahl}:{ConverttoChar(startBuchstabe + 3)}{startZahl + 1}"].Merge();
                    ws.Range[$"{ConverttoChar(startBuchstabe + 2)}{startZahl}:{ConverttoChar(startBuchstabe + 3)}{startZahl + 1}"].Interior.Color = Color.LightGray;
                    //ws.Range[$"{ConverttoChar(startBuchstabe + 2)}{startZahl}"].Value = name;
                    ErstelleDicLinks(name + maInfo, ws.Range[$"{ConverttoChar(startBuchstabe + 2)}{startZahl}"]);
                    startZahl++;

                }
                else
                {
                    ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}:{ConverttoChar(startBuchstabe + 1)}{startZahl}"].Merge();
                    ws.Range[$"{ConverttoChar(startBuchstabe)}{startZahl}"].Value = maInfo;
                    ws.Range[$"{ConverttoChar(startBuchstabe + 2)}{startZahl}:{ConverttoChar(startBuchstabe + 3)}{startZahl}"].Merge();
                    ws.Range[$"{ConverttoChar(startBuchstabe + 2)}{startZahl}"].Interior.Color = Color.LightGray;
                    if (counter == 0)
                    {
                        ws.Range[$"{ConverttoChar(startBuchstabe + 2)}{startZahl}"].Value = name;
                        counter++;
                    }
                    ErstelleDicLinks(name + maInfo, ws.Range[$"{ConverttoChar(startBuchstabe + 2)}{startZahl}"]);
                }

                startZahl++;
            }


        }


        private void CreateDropDown(string Cell, List<string> Items)
        {
            Excel.DropDowns xlDropDowns;
            Excel.DropDown xlDropDown;
            Range xlsRange;
            xlsRange = WS.Range[Cell];
            xlDropDowns = (Excel.DropDowns)(WS.DropDowns());
            xlDropDown = xlDropDowns.Add((double)xlsRange.Left, (double)xlsRange.Top, (double)xlsRange.Width, (double)xlsRange.Height, true);
            int index = 0;

            foreach (var item in Items)
            {
                xlDropDown.AddItem(item, index + 1);
            }
        }

        private string ConverttoChar(int number)
        {
            string result = "";
            while (number > 0)
            {
                number--;  // Verringern der Zahl, da der Index bei 0 beginnt
                result = (char)(number % 26 + 'A') + result; // Berechnen des Buchstabens und voranstellen
                number /= 26; // Aufteilen der Zahl für den nächsten Buchstaben
            }

            return result;
        }


        private void ErstelleDicLinks(string key, Range range)
        {
            if (KeyCells != null && (bool)!KeyCells?.ContainsKey(key))
            {
                KeyCells.Add(key, range);
            }
        }


    }
}