using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using EducServLib;
using BaseFormsLib;

namespace Priem
{
    public delegate void OnProfileInInObrazProgramInEntrySave(int ProfileId, int KCP);
    public partial class CardProfileInObrazProgramInEntry : BaseForm
    {
        private string Id;
        public event OnProfileInInObrazProgramInEntrySave OnSave;
        private int ProfileId
        {
            get { return ComboServ.GetComboIdInt(cbProfile).Value; }
            set { ComboServ.SetComboId(cbProfile, value); }
        }
        private int KCP
        {
            get 
            {
                int iRet = 0;
                int.TryParse(tbKCP.Text, out iRet);
                return iRet;
            }
            set { tbKCP.Text = value.ToString(); }
        }
        public CardProfileInObrazProgramInEntry(string sId, string LicenseProgramName, string ObrazProgramName)
        {
            InitializeComponent();
            tbLicenseProgramName.Text = LicenseProgramName;
            tbObrazProgramName.Text = ObrazProgramName;

            if (!string.IsNullOrEmpty(sId))
                Id = sId;

            this.MdiParent = MainClass.mainform;
            InitControls();
        }
        
        private void InitControls()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.SP_Profile.Select(x => new { x.Id, x.Name }).ToList().Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbProfile, src, true, false);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (OnSave != null)
                OnSave(ProfileId, KCP);

            this.Close();
        }
    }
}
