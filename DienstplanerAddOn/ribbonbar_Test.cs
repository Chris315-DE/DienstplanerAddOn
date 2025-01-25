using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Office = Microsoft.Office.Core;
using System.Web;
using DienstplanerAddOn.Lib;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using DienstplanerAddOn.Lib.Setup;



namespace DienstplanerAddOn
{
    public partial class Ribbon1
    {
        AddInManager Manager { get; set; }
        Workbook wb { get; set; }
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            KeyCellsManager manager = new KeyCellsManager();
            wb = Globals.ThisAddIn.GetActivWB();
            btnLoadMitarbeiter.Enabled = false;
            Manager = new AddInManager(wb);
        }



        private void button2_Click(object sender, RibbonControlEventArgs e)
        {

            Manager.SetupMitarbeiterListe(wb);
            var Button = (RibbonButton)sender;
            if (Button != null)
            {
                Button.Enabled = false;
                btnLoadMitarbeiter.Enabled = true;
            }
        }

        private void btnLoadMitarbeiter_Click(object sender, RibbonControlEventArgs e)
        {

            bool test = Manager.ErstelleMitarbeiterModel();
            if (!test)
            {
                MessageBox.Show("Bitte füllen sie die Tabelle MitarbeiterListe aus, bevor sie Setup2 starten");
            }
            else
            {
                Manager.ErstelleMitarbeiterTabelle(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["MitarbeiterListe"]);
                MessageBox.Show(Manager.AuswertungMA.MitarbeiterTypen[0].ToString());
            }


        }
        private void btn_DienstplanSetup_Click(object sender, RibbonControlEventArgs e)
        {

            Manager.ErstelleDienstplanSetup();
            var Button = (RibbonButton)sender;
            if (Button != null) 
            {
                Button.Enabled = false;
            }
        }
    }
}






