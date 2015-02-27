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
    public partial class CardApplication_InnerEntryInEntryPriorities : Form
    {
        private Guid _ApplicationId;
        private DataTable tblPrior;
        private int currSelectedRowId;
        public CardApplication_InnerEntryInEntryPriorities(Guid AppId)
        {
            InitializeComponent();
            _ApplicationId = AppId;

            this.MdiParent = MainClass.mainform;

            InitGrid();
            UpdateGrid();

            this.Focus();
        }

        private void InitGrid()
        {
            using (PriemEntities context = new PriemEntities())
            {
                Guid EntryId = context.Abiturient.Where(x => x.Id == _ApplicationId).Select(x => x.EntryId).FirstOrDefault();

                tbLicenseProgram.Text = context.Entry.Where(x => x.Id == EntryId).Select(x => x.SP_LicenseProgram.Code + " " + x.SP_LicenseProgram.Name).First();

                var baseData = context.InnerEntryInEntry.Where(x => x.EntryId == EntryId).Select(x => new
                {
                    x.Id,
                    ObrazProgramName = x.SP_ObrazProgram.Name,
                    //HasInnerProfiles = x.Count() > 1
                });

                var usedData = context.ApplicationDetails.Where(x => x.ApplicationId == _ApplicationId).Select(x => new
                {
                    x.InnerEntryInEntryId,
                    Priority = x.InnerEntryInEntryPriority
                });

                Dictionary<Guid, int> dicObrazPrograms = new Dictionary<Guid, int>();
                foreach (var usedProf in usedData)
                {
                    if (dicObrazPrograms.ContainsKey(usedProf.InnerEntryInEntryId))
                        continue;

                    int priorMin = usedData.Where(x => x.InnerEntryInEntryId == usedProf.InnerEntryInEntryId).Select(x => x.Priority).Min();

                    if (dicObrazPrograms.ContainsValue(priorMin) )
                        continue;
                    else
                        dicObrazPrograms.Add(usedProf.InnerEntryInEntryId, priorMin);
                }

                if (dicObrazPrograms.Count() < baseData.Count())
                    lblWarning.Visible = true;
                else
                    lblWarning.Visible = false;

                tblPrior = new DataTable();
                tblPrior.Columns.Add("Id", typeof(Guid));
                tblPrior.Columns.Add("Priority", typeof(int));
                tblPrior.Columns.Add("ObrazProgramName", typeof(string));
                tblPrior.Columns.Add("HasInnerProfilePriorityText", typeof(string));

                List<int> lst = new List<int>();
                for (int i = 1; i <= baseData.Count(); i++)
                    lst.Add(i);

                List<Guid> lstUsedObrazPrograms = new List<Guid>();
                lstUsedObrazPrograms.AddRange(dicObrazPrograms.Select(x => x.Key));
                foreach (int priority in lst)
                {
                    DataRow rw = tblPrior.NewRow();
                    Guid? Op = dicObrazPrograms.Where(x => x.Value == priority).Select(x => x.Key).FirstOrDefault();
                    if (Op == Guid.Empty)
                    {
                        Op = baseData.Where(x => !lstUsedObrazPrograms.Contains(x.Id)).Select(x => x.Id).FirstOrDefault();
                        if (Op != Guid.Empty)
                            lstUsedObrazPrograms.Add(Op.Value);
                        else
                            continue;
                    }

                    rw.SetField<Guid>("Id", Op.Value);
                    rw.SetField<string>("ObrazProgramName", baseData.Where(x=> x.Id == Op.Value).Select(x => x.ObrazProgramName).FirstOrDefault());
                    rw.SetField<int>("Priority", priority);
                    bool hasInnerProfiles = baseData.Where(x => x.Id == Op.Value).Count() > 1;
                    if (hasInnerProfiles)
                        rw.SetField<string>("HasInnerProfilePriorityText", "по профилям");
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
            dgv.Columns["ObrazProgramName"].HeaderText = "Обр. программа";
            dgv.Columns["ObrazProgramName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgv.Columns["HasInnerProfilePriorityText"].HeaderText = "";

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
            rw.SetField<string>("HasInnerProfilePriorityText", tblPrior.Rows[rwInd].Field<string>("HasInnerProfilePriorityText"));
            rw.SetField<string>("ObrazProgramName", tblPrior.Rows[rwInd].Field<string>("ObrazProgramName"));

            //insert row on new place
            tblPrior.Rows.InsertAt(rw, rwInd - 1);
            //delete from old place
            tblPrior.Rows.RemoveAt(rwInd + 1);

            for (int i = rwInd - 1; i < tblPrior.Rows.Count; i++)
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
            rw.SetField<string>("HasInnerProfilePriorityText", tblPrior.Rows[rwInd].Field<string>("HasInnerProfilePriorityText"));
            rw.SetField<string>("ObrazProgramName", tblPrior.Rows[rwInd].Field<string>("ObrazProgramName"));

            //insert row on new place
            tblPrior.Rows.InsertAt(rw, rwInd + 2);
            //delete from old place
            tblPrior.Rows.RemoveAt(rwInd);

            for (int i = rwInd; i < tblPrior.Rows.Count; i++)
            {
                tblPrior.Rows[i].SetField<int>("Priority", i + 1);
                if (tblPrior.Rows[i].Field<Guid>("Id") == NewApplicationId)
                    currSelectedRowId = i;
            }

            UpdateGrid();
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Сохранить изменения в приоритетах образовательных программ внутри конкурса?", "Изменение приоритетов", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                using (PriemEntities context = new PriemEntities())
                {
                    foreach (DataRow rw in tblPrior.Rows)
                    {
                        Guid ObrazProgramInEntryId = rw.Field<Guid>("Id");
                        context.Abiturient_UpdateObrazProgramInEntryPriority(ObrazProgramInEntryId, rw.Field<int>("Priority"), _ApplicationId);
                    }
                }

                lblWarning.Visible = false;
            }
        }

        private void dgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            if (string.IsNullOrEmpty(dgv["HasInnerProfilePriorityText", e.RowIndex].Value.ToString()))
                return;
        }
    }
}
