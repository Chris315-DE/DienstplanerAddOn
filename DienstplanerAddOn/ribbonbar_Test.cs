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
using Button = System.Windows.Forms.Button;


namespace DienstplanerAddOn
{
    public partial class Ribbon1
    {
        AddInManager Manager { get; set; }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            this.btnLoadMitarbeiter.Enabled = false;
        }



        private void button2_Click(object sender, RibbonControlEventArgs e)
        {



            Excel.Worksheet activesheet = Globals.ThisAddIn.Application.ActiveSheet;
            KeyCellsManager manager = new KeyCellsManager();
            Workbook wb = Globals.ThisAddIn.GetActivWB();
            Manager = new AddInManager(wb);

            var Button = (RibbonButton)sender;
            if (Button != null)
            {
                Button.Enabled = false;
                this.btnLoadMitarbeiter.Enabled = true;
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

                MessageBox.Show(Manager.AuswertungMA.MitarbeiterTypen[0].ToString());
            }

        }
    }
}
