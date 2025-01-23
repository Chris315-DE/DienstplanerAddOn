using DienstplanerAddOn.Lib.HelperClass;
using DienstplanerAddOn.Lib.Setup;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace DienstplanerAddOn.Lib
{
    public class AddInManager
    {
        Excel.Workbook aWb;
        public MitarbeiterListe mitarbeiterListe { get; set; }
        public Worksheet MitarbeiterSheet { get; set; }

        public AuswertungMitarbeiterListe AuswertungMA;


        public AddInManager(Workbook Wb)
        {
            aWb = Wb;
            SetupMitarbeiterListe(Wb);
           
          
        }


        private AddInManager SetupMitarbeiterListe(Workbook Wb)
        {
            aWb.Application.Worksheets.Add();
            Worksheet ws = aWb.Application.ActiveSheet as Worksheet;
            
            ws.Name = "MitarbeiterListe";
            MitarbeiterSheet = ws;
            mitarbeiterListe = new MitarbeiterListe(ws);

           


            return this;

        }


        public bool ErstelleMitarbeiterModel()
        {
            AuswertungMA = new AuswertungMitarbeiterListe(MitarbeiterSheet, mitarbeiterListe.tabelle);
            if (AuswertungMA.MitarbeiterTypen.Count > 0) 
            {
                return true;
            }
            return false;


        }










    }
}





