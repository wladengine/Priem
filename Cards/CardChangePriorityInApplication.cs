using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using EducServLib;


namespace Priem
{
    public delegate void UpdatePrirityHandler(int val);
    public partial class CardChangePriorityInApplication : Form
    {
        private DataTable tblPrior;
        private Guid _ApplicationId;
        public event UpdatePrirityHandler OnSave;
        private int currSelectedRowId = 0;

        public CardChangePriorityInApplication(Guid Id)
        {
            InitializeComponent();
            _ApplicationId = Id;

            this.MdiParent = MainClass.mainform;

            GetData();
            UpdateGrid();
        }

        private void GetData()
        {
            using (PriemEntities context = new PriemEntities())
            {
                Guid PersonId = context.Abiturient.Where(x => x.Id == _ApplicationId).Select(x => x.PersonId).DefaultIfEmpty(Guid.Empty).First();
                if (PersonId != Guid.Empty)
                {
                    var data = context.Abiturient
                        .Where(x => x.PersonId == PersonId && x.Entry.StudyLevel.LevelGroupId == MainClass.studyLevelGroupId && !x.IsGosLine && x.Entry.StudyBasisId == 1 && !x.BackDoc)
                        .Select(x => new
                        {
                            x.Id,
                            x.Priority,
                            LicenseProgram = x.Entry.SP_LicenseProgram.Code + " " + x.Entry.SP_LicenseProgram.Name,
                            ObrazProgram = x.Entry.StudyLevel.Acronym + "." + x.Entry.SP_ObrazProgram.Number + "." + MainClass.sPriemYear + " " + x.Entry.SP_ObrazProgram.Name,
                            Profile = x.Entry.SP_Profile.Name,
                            StudyForm = x.Entry.StudyForm.Name,
                        }).OrderBy(x => x.Priority).ToArray();

                    tblPrior = Util.ConvertToDataTable(data);

                    for (int i = 0; i < tblPrior.Rows.Count; i++)
                    {
                        if (tblPrior.Rows[i].Field<Guid>("Id") == _ApplicationId)
                            currSelectedRowId = i;
                    }
                }
            }
        }

        private void UpdateGrid()
        {
            dgv.DataSource = tblPrior;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.MultiSelect = false;
            dgv.Columns["Id"].Visible = false;
            dgv.Columns["LicenseProgram"].HeaderText = "Направление";
            dgv.Columns["LicenseProgram"].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dgv.Columns["ObrazProgram"].HeaderText = "Обр программа";
            dgv.Columns["ObrazProgram"].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dgv.Columns["Profile"].HeaderText = "Профиль";
            dgv.Columns["Profile"].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dgv.Columns["StudyForm"].HeaderText = "Форма обучения";
            dgv.Columns["Priority"].HeaderText = "Приоритет";
            dgv.Columns["Priority"].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

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

            Guid AppId = (Guid)dgv["Id", rwInd].Value;

            var rw = tblPrior.NewRow();
            rw.SetField<Guid>("Id", tblPrior.Rows[rwInd].Field<Guid>("Id"));
            rw.SetField<string>("LicenseProgram", tblPrior.Rows[rwInd].Field<string>("LicenseProgram"));
            rw.SetField<string>("ObrazProgram", tblPrior.Rows[rwInd].Field<string>("ObrazProgram"));
            rw.SetField<string>("Profile", tblPrior.Rows[rwInd].Field<string>("Profile"));
            rw.SetField<string>("StudyForm", tblPrior.Rows[rwInd].Field<string>("StudyForm"));

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

            Guid AppId = (Guid)dgv["Id", rwInd].Value;

            var rw = tblPrior.NewRow();
            rw.SetField<Guid>("Id", tblPrior.Rows[rwInd].Field<Guid>("Id"));
            rw.SetField<string>("LicenseProgram", tblPrior.Rows[rwInd].Field<string>("LicenseProgram"));
            rw.SetField<string>("ObrazProgram", tblPrior.Rows[rwInd].Field<string>("ObrazProgram"));
            rw.SetField<string>("Profile", tblPrior.Rows[rwInd].Field<string>("Profile"));
            rw.SetField<string>("StudyForm", tblPrior.Rows[rwInd].Field<string>("StudyForm"));
            
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
            if (MessageBox.Show("Сохранить изменения в приоритетах?", "Изменение приоритетов", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                int newPrior = 0;
                using (PriemEntities context = new PriemEntities())
                {
                    foreach (DataRow rw in tblPrior.Rows)
                    {
                        Guid AppId = rw.Field<Guid>("Id");
                        context.Abiturient_UpdatePriority(rw.Field<int>("Priority"), AppId);
                        if (AppId == _ApplicationId)
                            newPrior = rw.Field<int>("Priority");
                    }
                }

                //MainClass.DataRefresh();
                if (OnSave != null)
                    OnSave(newPrior);

                this.Close();
            }
        }
    }
}
