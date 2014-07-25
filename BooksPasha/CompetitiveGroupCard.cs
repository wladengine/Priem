using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Priem
{
    public partial class CompetitiveGroupCard : BookCard
    {
        public CompetitiveGroupCard(string id) : base(id)
        {
            InitializeComponent();
            InitControls();
        }

        protected override void ExtraInit()
        {
            this.MdiParent = MainClass.mainform;
            _tableName = "ed.CompetitiveGroup";
            base.ExtraInit();
        }

        private string CompGroupName
        {
            get { return tbName.Text.Trim(); }
            set { tbName.Text = value; }
        }

        private int? KCP
        {
            get 
            { 
                int retKCP = 0;
                if (!int.TryParse(tbKCP.Text.Trim(), out retKCP))
                    return null;
                return retKCP;
            }
            set 
            {
                tbKCP.Text = value.HasValue ? value.Value.ToString() : "";
            }
        }
        private void UpdateGrid()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ids = context.qEntryToCompetitiveGroup.Where(x => x.CompetitiveGroupId == GuidId).Select(x => x.EntryId).ToList();
                var src = context.qEntry.Where(x => ids.Contains(x.Id)).Select(x => new
                {
                    x.Id,
                    x.LicenseProgramCode,
                    x.LicenseProgramName,
                    x.ObrazProgramCrypt,
                    x.ObrazProgramName,
                    x.ProfileName,
                    x.StudyFormName,
                    x.StudyBasisName,
                    x.IsSecond,
                    x.IsReduced,
                    x.IsParallel
                }).ToList().Select(x => new
                {
                    x.Id,
                    LP = x.LicenseProgramCode + " " + x.LicenseProgramName,
                    OP = x.ObrazProgramCrypt + " " + x.ObrazProgramName + (x.IsReduced ? " (сокр)" : "") + (x.IsSecond ? " (для лиц с ВО)" : "") + (x.IsParallel ? " (параллел)" : ""),
                    x.ProfileName,
                    x.StudyFormName,
                    x.StudyBasisName
                }).ToList();
                
                dgvEntries.DataSource = src;
                dgvEntries.Columns["Id"].Visible = false;
            }
        }

        protected override void FillCard()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var data = context.CompetitiveGroup.Where(x => x.Id == GuidId).FirstOrDefault();

                if (data != null)
                {
                    CompGroupName = data.Name;
                    KCP = data.KCP;
                    UpdateGrid();
                }
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dgvEntries.SelectedCells.Count == 0)
                return;
            int rwInd = dgvEntries.SelectedCells[0].RowIndex;

            using (PriemEntities context = new PriemEntities())
            {
                Guid entryId = (Guid)dgvEntries["Id", rwInd].Value;
                context.EntryToCompetitiveGroup_Delete(entryId, GuidId);
            }

            UpdateGrid();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            var crd = new EntryForEntryInCompetitiveGroupCard();
            crd.OnOK += AddEntryToCompGroup;
            crd.Show();
        }

        private void AddEntryToCompGroup(Guid EntryId)
        {
            if (!GuidId.HasValue)
                return;

            using (PriemEntities context = new PriemEntities())
            {
                context.EntryToCompetitiveGroup_Insert(EntryId, GuidId);
            }
        }

        protected override void InsertRec(PriemEntities context, System.Data.Objects.ObjectParameter idParam)
        {
            context.CompetitiveGroup_Insert(CompGroupName, KCP, idParam);
            _Id = ((Guid)idParam.Value).ToString();
        }

        protected override void UpdateRec(PriemEntities context, Guid id)
        {
            context.CompetitiveGroup_update(CompGroupName, KCP, id);
        }
    }
}
