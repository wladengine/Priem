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
    public partial class PayDataEntryList : Form
    {
        public PayDataEntryList()
        {
            InitializeComponent();
            InitHandlers();
            FillComboStudyLevel();
            this.MdiParent = MainClass.mainform;
        }

        private void InitHandlers()
        {
            cbStudyLevel.SelectedIndexChanged += cbStudyLevel_SelectedIndexChanged;
            cbStudyForm.SelectedIndexChanged += cbStudyForm_SelectedIndexChanged;
            cbLicenseProgram.SelectedIndexChanged += cbLicenseProgram_SelectedIndexChanged;
            cbFaculty.SelectedIndexChanged += cbFaculty_SelectedIndexChanged;
            cbProrector.SelectedIndexChanged += cbProrector_SelectedIndexChanged;

            btnAdd.Click += btnAdd_Click;
            btnDelete.Click += btnDelete_Click;
            dgvList.CellDoubleClick += dgvList_CellDoubleClick;
        }
        private void NullHandlers()
        {
            cbStudyLevel.SelectedIndexChanged -= cbStudyLevel_SelectedIndexChanged;
            cbStudyForm.SelectedIndexChanged -= cbStudyForm_SelectedIndexChanged;
            cbLicenseProgram.SelectedIndexChanged -= cbLicenseProgram_SelectedIndexChanged;
            cbFaculty.SelectedIndexChanged -= cbFaculty_SelectedIndexChanged;
            cbProrector.SelectedIndexChanged -= cbProrector_SelectedIndexChanged;

            btnAdd.Click -= btnAdd_Click;
            btnDelete.Click -= btnDelete_Click;
            dgvList.CellDoubleClick -= dgvList_CellDoubleClick;
        }

        public int? StudyLevelId
        {
            get { return ComboServ.GetComboIdInt(cbStudyLevel); }
        }
        public int? FacultyId
        {
            get { return ComboServ.GetComboIdInt(cbFaculty); }
        }
        public int? StudyFormId
        {
            get { return ComboServ.GetComboIdInt(cbStudyForm); }
        }
        public int? LicenseProgramId
        {
            get { return ComboServ.GetComboIdInt(cbLicenseProgram); }
        }
        public int? ProrectorId
        {
            get { return ComboServ.GetComboIdInt(cbProrector); }
        }

        void cbProrector_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }
        void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboLicenseProgram();
        }
        void cbLicenseProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboProrector();
            FillGrid();
        }
        void cbStudyForm_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboFaculty();
        }
        void cbStudyLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboStudyForm();
        }

        void btnDelete_Click(object sender, EventArgs e)
        {
            //int rwInd 
        }

        void btnAdd_Click(object sender, EventArgs e)
        {
            var crd = new PayDataEntryCard(StudyLevelId, StudyFormId, FacultyId, LicenseProgramId);
            crd.ToUpdateList += FillGrid;
            crd.Show();
        }

        void dgvList_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;
            var crd = new PayDataEntryCard((Guid)dgvList["EntryId", e.RowIndex].Value);
            crd.ToUpdateList += FillGrid;
            crd.Show();
        }

        private void FillComboStudyLevel()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.qEntry.Select(x => new { x.StudyLevelId, x.StudyLevelName }).Distinct().ToList()
                    .Select(x => new KeyValuePair<string, string>(x.StudyLevelId.ToString(), x.StudyLevelName)).ToList();

                ComboServ.FillCombo(cbStudyLevel, src, false, true);
            }
        }
        private void FillComboStudyForm()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.qEntry.Where(x => StudyLevelId.HasValue ? x.StudyLevelId == StudyLevelId : true).Select(x => new { Id = x.StudyFormId, Name = x.StudyFormName }).Distinct().ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();

                ComboServ.FillCombo(cbStudyForm, src, false, true);
            }
        }
        private void FillComboFaculty()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.qEntry.Where(x => (StudyLevelId.HasValue ? x.StudyLevelId == StudyLevelId : true) && (StudyFormId.HasValue ? x.StudyFormId == StudyFormId : true))
                    .Select(x => new { Id = x.FacultyId, Name = x.FacultyName }).Distinct().ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();

                ComboServ.FillCombo(cbFaculty, src, false, true);
            }
        }
        private void FillComboLicenseProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.qEntry.Where(x => (StudyLevelId.HasValue ? x.StudyLevelId == StudyLevelId : true) && (StudyFormId.HasValue ? x.StudyFormId == StudyFormId : true) && (FacultyId.HasValue ? x.FacultyId == FacultyId : true)) 
                    .Select(x => new { Id = x.LicenseProgramId, Name = x.LicenseProgramName, x.LicenseProgramCode }).Distinct().ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.LicenseProgramCode + " " + x.Name)).ToList();

                ComboServ.FillCombo(cbLicenseProgram, src, false, true);
            }
        }
        private void FillComboProrector()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.Prorektor.Select(x => new { Id = x.Id, Name = x.Name}).Distinct().ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();

                ComboServ.FillCombo(cbProrector, src, false, true);
            }
        }

        private void FillGrid()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var data = (from x in context.PayDataEntry
                            join Faculty in context.SP_Faculty on x.Entry.FacultyId equals Faculty.Id
                            where (StudyFormId.HasValue ? x.Entry.StudyFormId == StudyFormId : true)
                            && (StudyLevelId.HasValue ? x.Entry.StudyLevelId == StudyLevelId : true)
                            && (FacultyId.HasValue ? x.Entry.FacultyId == FacultyId : true)
                            && (LicenseProgramId.HasValue ? x.Entry.LicenseProgramId == LicenseProgramId : true)
                            select new
                            {
                                x.EntryId,
                                Faculty = Faculty.Name,
                                LicenseProgramCode = x.Entry.SP_LicenseProgram.Code,
                                LicenseProgramName = x.Entry.SP_LicenseProgram.Name,
                                ObrazProgramCrypt = x.Entry.StudyLevel.Acronym + "." + x.Entry.SP_ObrazProgram.Number + "." + MainClass.sPriemYear,
                                ObrazProgramName = x.Entry.SP_ObrazProgram.Name,
                                ProfileName = x.Entry.SP_Profile.Name,
                                StudyForm = x.Entry.StudyForm.Name,
                                Prorektor = x.Prorektor.Name
                            }).ToList()
                            .Select(x =>
                               new
                               {
                                   x.EntryId,
                                   LP = x.LicenseProgramCode + " " + x.LicenseProgramName,
                                   OP = x.ObrazProgramCrypt + " " + x.ObrazProgramName,
                                   x.ProfileName,
                                   x.StudyForm,
                                   x.Prorektor
                               }).OrderBy(x => x.LP).ToList();

                dgvList.DataSource = data;
                dgvList.Columns["EntryId"].Visible = false;
            }

        }

        private void btnDelete_Click_1(object sender, EventArgs e)
        {
            if (dgvList.SelectedCells.Count == 0)
                return;
            int rwInd = dgvList.SelectedCells[0].RowIndex;

            if (MessageBox.Show("Удалить запись?", "Внимание", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.No)
                return;

            Guid EntryId = (Guid)dgvList["EntryId", rwInd].Value;

            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    context.PayDataEntry_Delete(EntryId);
                }

                FillGrid();
            }
            catch (Exception ex)
            {
                WinFormsServ.Error(ex);
            }
        }
    }
}
