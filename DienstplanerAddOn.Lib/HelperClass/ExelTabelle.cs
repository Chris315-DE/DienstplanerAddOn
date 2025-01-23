using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DienstplanerAddOn.Lib.HelperClass
{
    public class ExelTabelle
    {

        public List<string> Headders;
        public int Cols;

        public int[] Rows;
        public int[] Ccols;




        public ExelTabelle(string name, List<string> Headders, int[] RowLenght, int[] ColHeight)
        {
            Name = name;
            Felder = new List<Cells>();
            Cols = ColHeight.Length;
            Ccols = ColHeight;
            Rows = RowLenght;
            this.Headders = Headders;

        }





        private int count = 0;
        public string Name { get; set; }

        public List<Cells> Felder { get; set; }


        /// <summary>
        /// Gibt die Celle wieder die dem Key oder der Gegsuchten Adresse Entsprechen
        /// 
        /// </summary>
        /// <param name="key">or Adress</param>
        /// <returns> <see cref="Cells"/>
        /// <see cref="InvalidOperationException"/>
        /// </returns>
        public Cells this[string key]
        {
            get
            {
                foreach (var cell in Felder)
                {
                    if (cell.Key == key)
                    {
                        return cell;
                    }

                    if (cell.Adress == key)
                        return cell;
                }
                throw new InvalidOperationException();
            }

        }

        public void AddCell(Range range, string key)
        {
            if (Felder == null)
                Felder = new List<Cells>();

            string adress = ConvertToExcelCell(range.Column, range.Row);

            Cells cells = new Cells(adress, $"{key}:{count}");



            while (Felder.Contains(cells))
            {
                count++;
                cells.Key = $"{key}:{count}";

            }

            count = 0;


            Felder.Add(cells);

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
    }

}
