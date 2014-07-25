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
    public partial class CompetitiveGroupList : Form
    {
        public CompetitiveGroupList()
        {
            InitializeComponent();
            this.MdiParent = MainClass.mainform;
            FillComboStudyLevel();
        }

        private void FillComboStudyLevel()
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<KeyValuePair<string, string>> lst = context.qEntry.Select(x => new { x.StudyLevelId, x.StudyLevelName }).Distinct().ToList()
                    .Select(x => new KeyValuePair<string, string>(x.StudyLevelId.ToString(), x.StudyLevelName)).ToList();

                ComboServ.FillCombo(cbStudyLevel, lst, true, false);
            }
        }
        public int? StudyLevelId
        {
            get { return ComboServ.GetComboIdInt(cbStudyLevel); }
            set { ComboServ.SetComboId(cbStudyLevel, value); }
        }

        private void FillGrid()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src =
                    (from CompGroup in context.CompetitiveGroup
                     join EntToCompGroup in context.qEntryToCompetitiveGroup on CompGroup.Id equals EntToCompGroup.CompetitiveGroupId into EntToCompGroup2
                     from EntToCompGroup in EntToCompGroup2.DefaultIfEmpty()
                     join E in context.Entry on EntToCompGroup.EntryId equals E.Id into E2
                     from E in E2.DefaultIfEmpty()
                     where (StudyLevelId.HasValue ? E.StudyLevelId == StudyLevelId : E.StudyLevelId == null)
                     select new
                     {
                         CompGroup.Id,
                         CompGroup.Name
                     }).Distinct().ToList();

                dgvList.DataSource = src;
                dgvList.Columns["Id"].Visible = false;
            }
        }

        private void cbStudyLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillGrid();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            new CompetitiveGroupCard(null).Show();
        }

        private void dgvList_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;
            string id = dgvList["Id", e.RowIndex].Value.ToString();
            new CompetitiveGroupCard(id).Show();
        }
    }
}
