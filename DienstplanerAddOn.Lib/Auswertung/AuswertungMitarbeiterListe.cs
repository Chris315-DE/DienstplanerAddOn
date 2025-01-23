using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using DienstplanerAddOn.Lib.Verwaltung;
using DienstplanerAddOn.Lib.HelperClass;
using System.Windows.Forms;
using System.Diagnostics;
namespace DienstplanerAddOn.Lib
{
    public class AuswertungMitarbeiterListe
    {
        Worksheet WS;
        ExelTabelle Tabelle;
        public List<MitarbeiterTyp> MitarbeiterTypen { get; set; }

        bool isInit = true;

        public AuswertungMitarbeiterListe(Worksheet activeWorksheet, ExelTabelle tabelle)
        {

            List<string> values = new List<string>();

            WS = activeWorksheet;
            Tabelle = tabelle;
            MitarbeiterTypen = new List<MitarbeiterTyp>();
            if (tabelle == null || WS == null)
            {
                MessageBox.Show("Fehler");
                isInit = false;
                return;
            }

            for (int i = 0; i < tabelle.Cols-1; i++)
            {
                foreach (var header in tabelle.Headders)
                {
                    string key = $"{header}:{i}";
                    Range r = WS.Range[tabelle[key].Adress];
                    if (r.Value == null)
                        continue;
                    if(header != tabelle.Headders[0])
                    {
                        int a;
                        if(!int.TryParse(r.Value.ToString(),out a)){
                            MessageBox.Show($"Fehler in der Zelle:{ConvertToExcelCell(r.Column, r.Row)}");
                            isInit = false;
                            break;
                        }
                    }


                    values.Add(r.Value.ToString());



                }



                if (values.Count == 7)
                {
                    string name = values[0];

                    int beproschicht = int.Parse(values[1]);
                    int beproTag = int.Parse(values[2]);
                    int MaxAmStück = int.Parse(values[3]);
                    int Frei = int.Parse(values[4]);
                    int MaxProMon = int.Parse(values[5]);
                    int MinProMon = int.Parse(values[6]);

                    MitarbeiterTypen.Add(new MitarbeiterTyp(name, beproschicht, beproTag, MaxAmStück, Frei, MaxProMon, MinProMon));

                    values.Clear();
                }
            }
          

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


    }
}
