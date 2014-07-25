using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using BaseFormsLib;
using EducServLib;

namespace Priem
{
    public partial class VuzNameCard : BookCardInt
    {
        private int? RegionId
        {
            get { return ComboServ.GetComboIdInt(cbRegion); }
            set { ComboServ.SetComboId(cbRegion, value); }
        }
        private string EntityName
        {
            get { return tbName.Text.Trim(); }
            set { tbName.Text = value; }
        }

        public VuzNameCard(string id) : base(id)
        {
            _tableName = "ed.SchoolNames";
            InitializeComponent();
            InitControls();
        }

        protected override void ExtraInit()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.Region.Select(x => new { x.Id, x.Name, x.RegionNumber, IsSpb = x.RegionNumber == "78" ? 1 : 0 }).ToList()
                    .Select(x => new { x.Id, x.IsSpb, x.Name, x.RegionNumber, HasNum = string.IsNullOrEmpty(x.RegionNumber) ? 0 : 1 })
                    .OrderByDescending(x => x.IsSpb).ThenByDescending(x => x.HasNum).ThenBy(x => x.RegionNumber)
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), (x.HasNum == 1 ? "[" + x.RegionNumber + "] " : "") + x.Name)).ToList();
                ComboServ.FillCombo(cbRegion, src, true, false);
            }
        }

        protected override void FillCard()
        {
            if (string.IsNullOrEmpty(_Id))
                return;

            using (PriemEntities context = new PriemEntities())
            {
                var ent = context.SchoolNames.Where(x => x.Id == IntId).Select(x => new { x.RegionId, x.Name }).FirstOrDefault();
                EntityName = ent.Name;
                RegionId = ent.RegionId;
            }
        }

        protected override void InsertRec(PriemEntities context, System.Data.Objects.ObjectParameter idParam)
        {
            context.SchoolNames_Insert(RegionId, 4, EntityName, idParam);
        }
        protected override void UpdateRec(PriemEntities context, int id)
        {
            context.SchoolNames_Update(RegionId, 4, EntityName, id);
        }

        protected override void CloseCardAfterSave()
        {
            this.Close();
        }
    }
}
