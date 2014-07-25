using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Priem.BooksPasha
{
    public partial class ParametersForm : Form
    {
        public ParametersForm()
        {
            InitializeComponent();

        }

        private void InitControls()
        {
            this.MdiParent = MainClass.mainform;
            ExtraInit();
            FillValues();
        }

        private void ExtraInit()
        {
            
        }

        private void FillValues()
        {
            using (PriemEntities context = new PriemEntities())
            {
                Dictionary<string, string> dic = context.C_AppSettings.Select(x => new { x.ParamKey, x.ParamValue }).ToList().ToDictionary(x => x.ParamKey, y => y.ParamValue);

                if (dic.ContainsKey("PriemYear"))
                    tbPriemYear.Text = dic["PriemYear"];

                if (dic.ContainsKey("b1kCheckProtocolsEnabled"))
                    chb1kCheckProtocolsEnabled.Checked = bool.Parse(dic["b1kCheckProtocolsEnabled"]);

                if (dic.ContainsKey("bMagCheckProtocolsEnabled"))
                    chbMagCheckProtocolsEnabled.Checked = bool.Parse(dic["bMagCheckProtocolsEnabled"]);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            //Год приёма
            int iVal;
            int.TryParse(MainClass.sPriemYear, out iVal);
            MainClass.iPriemYear = iVal;
            MainClass.sPriemYear = tbPriemYear.Text;
            //Флаги проверки протоколов о допуске
            MainClass.b1kCheckProtocolsEnabled = chb1kCheckProtocolsEnabled.Checked;
            MainClass.bMagCheckProtocolsEnabled = chbMagCheckProtocolsEnabled.Checked;
        }
    }
}
