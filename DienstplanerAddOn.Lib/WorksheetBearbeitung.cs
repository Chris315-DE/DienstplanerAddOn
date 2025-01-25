using DienstplanerAddOn.Lib.HelperClass;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace DienstplanerAddOn.Lib
{
    public abstract class WorksheetBearbeitung
    {
        Worksheet ws;

        internal int Buchstabe = 1;
        internal int Zahl = 1;

        Dictionary<string, Range> Keys;

        public WorksheetBearbeitung(Worksheet worksheet)
        {
            ws = worksheet;

        }
        /// <summary>
        /// Convertiert 2 Integer Werte zu einer Excel Cell
        /// </summary>
        /// <param name="buchstabe">1 = A , 2 = B, 3 = C usw,</param>
        /// <param name="zahl">Die Cell Number</param>
        /// <returns></returns>
        internal string ConvertToExcelCell(int buchstabe, int zahl)
        {
            string result = "";
            while (buchstabe > 0)
            {
                buchstabe--;  // Verringern der Zahl, da der Index bei 0 beginnt
                result = (char)(buchstabe % 26 + 'A') + result; // Berechnen des Buchstabens und voranstellen
                buchstabe /= 26; // Aufteilen der Zahl für den nächsten Buchstaben
            }

            result += zahl;

            return result;
        }
        /// <summary>
        /// Macht das Selbe wie 
        /// <see cref="ConvertToExcelCell(int, int)" />
        /// Nur für eine Range zb A1:N2
        /// </summary>
        /// <param name="buchstabe">
        /// <see cref="ConvertToExcelCell(int, int)"/></param>
        /// <param name="zahl">
        /// <see cref="ConvertToExcelCell(int, int)"/></param>
        /// <param name="buchstabe2">
        /// Nimmt die Werte für das ende der Rang an 
        /// Gleiches Spiel wie bei <see cref="ConvertToExcelCell(int, int)"/>
        /// </param>
        /// <param name="zahl2">
        /// Die Cell Number</param>
        /// <returns></returns>
        internal string ConvertToExcelCell(int buchstabe, int zahl, int buchstabe2, int zahl2)
        {
            string result = "";
            while (buchstabe > 0)
            {
                buchstabe--;  // Verringern der Zahl, da der Index bei 0 beginnt
                result = (char)(buchstabe % 26 + 'A') + result; // Berechnen des Buchstabens und voranstellen
                buchstabe /= 26; // Aufteilen der Zahl für den nächsten Buchstaben
            }

            result += zahl;

            result += ":";
            string result2 = "";
            while (buchstabe2 > 0)
            {
                buchstabe2--; //Verringern der Zahl, da der Index bei 0 beginnt
                result2 = (char)(buchstabe2 % 26 + 'A') + result2;
                buchstabe2 /= 26;
            }

            result += result2;
            result += zahl2;

            return result;
        }


        internal void SetDic(ref Dictionary<string, Range> dic)
        {
            Keys = dic;
        }

        public void CreateHeader(string headertext, string descript)
        {
            var EditCell = ws.Range[ConvertToExcelCell(Buchstabe, Zahl, Buchstabe + 14, Zahl + 1)];
            EditCell.Merge();
            EditCell.Value = headertext;
            EditCell.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            EditCell.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
            EditCell.Cells.Font.Bold = true;
            EditCell = ws.Range[ConvertToExcelCell(Buchstabe, Zahl + 2, Buchstabe + 14, Zahl + 4)];
            EditCell.Merge();
            EditCell.Value = descript;
            EditCell.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            EditCell.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
            EditCell.Cells.Font.Bold = true;
            Zahl += 6;

        }

        internal abstract void ErstelleKeyDicEintrag(string key, Range value);


        internal abstract Range GetRange(string key);



        /// <summary>
        /// Gibt den Wert der Excel Celle als String zurück
        /// 
        /// </summary>
        /// <param name="ws">Das Worksheet das ausgelesen werden soll</param>
        /// <param name="buchstabe"><see cref="ConvertToExcelCell(int, int)"/> </param>
        /// <param name="zahl"><see cref="ConvertToExcelCell(int, int)"/></param>
        /// <returns></returns>

        internal string GetValue(Worksheet ws, int buchstabe, int zahl)
        {
            return ws.Range[ConvertToExcelCell(buchstabe, zahl)].Value;
        }

        internal int CreateTabelle(Worksheet ws, int startrow, int startcol, Color color, Dictionary<string, Range> keyvalue, ExelTabelle tabelle)
        {
            if (ws == null) return 0;
            int buchstabe = startrow;
            int zahl = startcol;
            int counter = 0;
            int row = 0;
            int col = 0;
            int debugrounds = 0;
            Range r;


            for (int i = 0; i <= tabelle.Cols; i++)
            {


                foreach (var rrow in tabelle.Rows)
                {

                    r = ws.Range[ConvertToExcelCell(buchstabe, zahl, buchstabe + rrow - 1, zahl + 1 - 1)];
                    r.Merge();

                    if (counter < tabelle.Headders.Count)
                    {
                        r.Value = tabelle.Headders[counter];
                        r.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        r.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        counter++;

                    }
                    else
                    {
                        r.Interior.Color = color;
                        r.Cells.BorderAround2();
                        ErstelleKeyDicEintrag($"{tabelle.Name}:{col}{row}", r);
                        tabelle.AddCell(r, tabelle.Headders[row]);

                        if(tabelle is FormulaTabelle)
                        {
                            var tab = (FormulaTabelle)tabelle;
                            if(tab.Position == FormulaPosition.Right)
                            {
                                if (tab.Headders[row] == tab.Headders.Last())
                                {
                                    r.Formula = tab.FormelValues[col];
                                }

                            }

                        }



                    }

                    buchstabe += rrow;
                    row++;

                }
                zahl += 1;
                buchstabe = startrow;
                col++;
                row = 0;
                debugrounds++;
            }





            return zahl;
        }

        internal string createFormula(List<string> namelist, string baseformula)
        {
            StringBuilder sb = new StringBuilder();

            //="IF(OR(PARAM),\"Gültig\",\"Ungültig\")"

            foreach (var t in namelist)
            {
                sb.Append($"Param0=\"{t}\"");
                if (t != namelist.Last())
                {
                    sb.Append(",");
                }
            }


            var back = baseformula.Replace("PARAM", sb.ToString());

            return back;

        }

    }
}
