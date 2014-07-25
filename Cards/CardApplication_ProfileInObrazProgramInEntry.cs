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
    public partial class CardApplication_ProfileInObrazProgramInEntry : Form
    {
        private Guid _ApplicationId;
        private Guid _ObrazProgramInEntryId;
        private DataTable tblPrior;
        private int currSelectedRowId;
        public CardApplication_ProfileInObrazProgramInEntry(Guid ApplicationId, Guid ObrazProgramInEntryId)
        {
            InitializeComponent();
            _ApplicationId = ApplicationId;
            _ObrazProgramInEntryId = ObrazProgramInEntryId;

            this.MdiParent = MainClass.mainform;

            InitGrid();
            UpdateGrid();
        }

        private void InitGrid()
        {
            using (PriemEntities context = new PriemEntities())
            {
                Guid EntryId = context.Abiturient.Where(x => x.Id == _ApplicationId).Select(x => x.EntryId).FirstOrDefault();

                var Entry = context.Entry.Where(x => x.Id == EntryId).FirstOrDefault();
                if (Entry == null)
                    return;

                tbLicenseProgram.Text = Entry.SP_LicenseProgram.Code + " " + Entry.SP_LicenseProgram.Name;
                tbObrazProgram.Text = context.ObrazProgramInEntry.Where(x => x.Id == _ObrazProgramInEntryId)
                    .Select(x => x.SP_ObrazProgram.SP_LicenseProgram.StudyLevel.Acronym + "." + x.SP_ObrazProgram.Number + "." + MainClass.sPriemYear + " " + x.SP_ObrazProgram.Name).First();

                var baseData = context.ProfileInObrazProgramInEntry.Where(x => x.ObrazProgramInEntryId == _ObrazProgramInEntryId).Select(x => new
                {
                    x.Id,
                   ProfileName = x.SP_Profile.Name,
                }).ToList();

                var usedData = context.ApplicationDetails.Where(x => x.ApplicationId == _ApplicationId && x.ObrazProgramInEntryId == _ObrazProgramInEntryId).Select(x => new
                {
                    x.ProfileInObrazProgramInEntryId,
                    Priority = x.ProfileInObrazProgramInEntryPriority
                }).ToList();

                Dictionary<Guid, int> dicProfiles = new Dictionary<Guid, int>();
                foreach (var usedProf in usedData)
                {
                    if (!usedProf.ProfileInObrazProgramInEntryId.HasValue)
                        continue;
                    if (dicProfiles.ContainsKey(usedProf.ProfileInObrazProgramInEntryId.Value))
                        continue;

                    int priorMin = usedData.Where(x => x.ProfileInObrazProgramInEntryId == usedProf.ProfileInObrazProgramInEntryId).Select(x => x.Priority ?? int.MaxValue).Min();

                    if (dicProfiles.ContainsValue(priorMin) || priorMin == int.MaxValue)
                        continue;
                    else
                        dicProfiles.Add(usedProf.ProfileInObrazProgramInEntryId.Value, priorMin);
                }

                if (dicProfiles.Count() < baseData.Count())
                    lblWarning.Visible = true;
                else
                    lblWarning.Visible = false;

                tblPrior = new DataTable();
                tblPrior.Columns.Add("Id", typeof(Guid));
                tblPrior.Columns.Add("Priority", typeof(int));
                tblPrior.Columns.Add("ProfileName", typeof(string));

                List<int> lst = new List<int>();
                for (int i = 1; i <= baseData.Count(); i++)
                    lst.Add(i);

                

                List<Guid> lstUsedProfiles = new List<Guid>();
                lstUsedProfiles.AddRange(dicProfiles.Keys);
                foreach (int priority in lst)
                {
                    DataRow rw = tblPrior.NewRow();
                    Guid? Prof = dicProfiles.Where(x => x.Value == priority).Select(x => x.Key).FirstOrDefault();
                    if (Prof == null || Prof == Guid.Empty)
                    {
                        Prof = baseData.Where(x => !lstUsedProfiles.Contains(x.Id)).Select(x => x.Id).First();
                        lstUsedProfiles.Add(Prof.Value);
                    }

                    rw.SetField<Guid>("Id", Prof.Value);
                    rw.SetField<string>("ProfileName", baseData.Where(x => x.Id == Prof.Value).Select(x => x.ProfileName).FirstOrDefault());
                    rw.SetField<int>("Priority", priority);
                    tblPrior.Rows.Add(rw);
                }
            }
        }
        private void UpdateGrid()
        {
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.ClearSelection();
            dgv.DataSource = tblPrior;
            dgv.Columns["Id"].Visible = false;
            dgv.Columns["Priority"].HeaderText = "Приор.";
            dgv.Columns["Priority"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
            dgv.Columns["ProfileName"].HeaderText = "Профиль";
            dgv.Columns["ProfileName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

            dgv.Rows[currSelectedRowId].Selected = true;
        }

        private void btnUp_Click(object sender, EventArgs e)
        {
            if (dgv.SelectedCells.Count == 0)
                return;

            //находим выбранную конкурсную позицию, которую нужно поднять
            int rwInd = dgv.SelectedCells[0].RowIndex;
            Guid NewApplicationId = (Guid)dgv["Id", rwInd].Value;
            if (rwInd == 0)
                return;

            Guid OPInEntryId = (Guid)dgv["Id", rwInd].Value;

            var rw = tblPrior.NewRow();
            rw.SetField<Guid>("Id", tblPrior.Rows[rwInd].Field<Guid>("Id"));
            rw.SetField<string>("ProfileName", tblPrior.Rows[rwInd].Field<string>("ProfileName"));

            //insert row on new place
            tblPrior.Rows.InsertAt(rw, rwInd - 1);
            //delete from old place
            tblPrior.Rows.RemoveAt(rwInd + 1);

            for (int i = 0; i < tblPrior.Rows.Count; i++)
            {
                tblPrior.Rows[i].SetField<int>("Priority", i + 1);
                if (tblPrior.Rows[i].Field<Guid>("Id") == NewApplicationId)
                    currSelectedRowId = i;
            }

            UpdateGrid();
        }
        private void btnDown_Click(object sender, EventArgs e)
        {
            if (dgv.SelectedCells.Count == 0)
                return;

            //находим выбранную конкурсную позицию, которую нужно опустить
            int rwInd = dgv.SelectedCells[0].RowIndex;
            Guid NewApplicationId = (Guid)dgv["Id", rwInd].Value;

            if (rwInd == dgv.Rows.Count)
                return;

            Guid OPInEntryId = (Guid)dgv["Id", rwInd].Value;

            var rw = tblPrior.NewRow();
            rw.SetField<Guid>("Id", tblPrior.Rows[rwInd].Field<Guid>("Id"));
            rw.SetField<string>("ProfileName", tblPrior.Rows[rwInd].Field<string>("ProfileName"));

            //insert row on new place
            tblPrior.Rows.InsertAt(rw, rwInd + 2);
            //delete from old place
            tblPrior.Rows.RemoveAt(rwInd);

            for (int i = 0; i < tblPrior.Rows.Count; i++)
            {
                tblPrior.Rows[i].SetField<int>("Priority", i + 1);
                if (tblPrior.Rows[i].Field<Guid>("Id") == NewApplicationId)
                    currSelectedRowId = i;
            }

            UpdateGrid();
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Сохранить изменения в приоритетах профилей в образовательных программах внутри конкурса?", "Изменение приоритетов", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                using (PriemEntities context = new PriemEntities())
                {
                    foreach (DataRow rw in tblPrior.Rows)
                    {
                        Guid ProfileInObrazProgramInEntryId = rw.Field<Guid>("Id");
                        context.Abiturient_UpdateProfileInObrazProgramInEntryPriority(ProfileInObrazProgramInEntryId, rw.Field<int>("Priority"), _ApplicationId);
                    }
                }

                lblWarning.Visible = false;

                this.Close();
            }
        }
    }
}
