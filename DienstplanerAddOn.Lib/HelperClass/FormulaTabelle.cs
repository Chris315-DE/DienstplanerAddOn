using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DienstplanerAddOn.Lib.HelperClass
{
    public class FormulaTabelle : ExelTabelle
    {

        public FormulaPosition Position { get; }
        public string Formula;
        private int _buchstabe;
        private int _zahl;
        private List<int> _paramcells;
        public List<string> FormelValues { get; private set; }

        public FormulaTabelle(string name, List<string> Headders, int[] RowLenght, int cols, string formula, FormulaPosition pos, int buchstabe, int zahl, List<int> Paramcells) : base(name, Headders, RowLenght, cols)
        {
            Position = pos;
            Formula = formula;
            FormelValues = new List<string>();
            _buchstabe = buchstabe;
            _zahl = zahl;
            _paramcells = Paramcells;

            generateFormelValues();

            int a = 1;
        }

        private void generateFormelValues()
        {
            string newFormula = string.Empty;
            int offset = 0;
            switch (Position)
            {
                case FormulaPosition.Right:

                    foreach (int co in Rows)
                    {
                        offset += co;
                    }

                    for (int j = 0; j <= Cols; j++)
                    {


                        for (int i = 0; i < _paramcells.Count; i++)
                        {



                            newFormula = Formula.Replace($"Param{i}", $"{ConvertToExcelCell(_buchstabe + _paramcells[i], _zahl + j)}");
                        }
                        FormelValues.Add(newFormula);
                    }

                    break;
                case FormulaPosition.Bottom:
                    break;

            }
        }



    }





  
}
