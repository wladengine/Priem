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
    public delegate void OnOkBtnCall(Guid id);
    public partial class EntryForEntryInCompetitiveGroupCard : Form
    {
        public event OnOkBtnCall OnOK;
        public EntryForEntryInCompetitiveGroupCard()
        {
            InitializeComponent();
            this.MdiParent = MainClass.mainform;
        }

        private void FillComboStudyLevel()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.qEntry.Select(x => new { x.StudyLevelId, x.StudyLevelName }).Distinct().ToList()
                    .Select(x => new KeyValuePair<string, string>(x.StudyLevelId.ToString(), x.StudyLevelName)).ToList();

                ComboServ.FillCombo(cbStudyLevel, src, false, false);
            }
        }
        private void FillComboStudyForm()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.qEntry.Where(x => true && (StudyLevelId.HasValue ? x.StudyLevelId == StudyLevelId : true))
                    .Select(x => new { Id = x.StudyFormId, Name = x.StudyFormName }).Distinct().ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();

                ComboServ.FillCombo(cbStudyForm, src, false, false);
            }
        }
        private void FillComboStudyBasis()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.qEntry.Where(x => true && (StudyLevelId.HasValue ? x.StudyLevelId == StudyLevelId : true))
                    .Select(x => new { Id = x.StudyBasisId, Name = x.StudyBasisName }).Distinct().ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();

                ComboServ.FillCombo(cbStudyBasis, src, false, false);
            }
        }
        private void FillComboFaculty()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.qEntry.Where(x => (StudyLevelId.HasValue ? x.StudyLevelId == StudyLevelId : true) && (StudyFormId.HasValue ? x.StudyFormId == StudyFormId : true))
                    .Select(x => new { Id = x.FacultyId, Name = x.FacultyName }).Distinct().ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();

                ComboServ.FillCombo(cbFaculty, src, false, false);
            }
        }
        private void FillComboLicenseProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.qEntry.Where(x => (StudyLevelId.HasValue ? x.StudyLevelId == StudyLevelId : true) && (StudyFormId.HasValue ? x.StudyFormId == StudyFormId : true) && (FacultyId.HasValue ? x.FacultyId == FacultyId : true))
                    .Select(x => new { Id = x.LicenseProgramId, Name = x.LicenseProgramName, x.LicenseProgramCode }).Distinct().ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.LicenseProgramCode + " " + x.Name)).ToList();

                ComboServ.FillCombo(cbLicenseProgram, src, false, false);
            }
        }
        private void FillComboObrazProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.qEntry.Where(x => (StudyLevelId.HasValue ? x.StudyLevelId == StudyLevelId : true)
                    && (StudyFormId.HasValue ? x.StudyFormId == StudyFormId : true)
                    && (LicenseProgramId.HasValue ? x.LicenseProgramId == LicenseProgramId : true)
                    && (FacultyId.HasValue ? x.FacultyId == FacultyId : true))
                    .Select(x => new { Id = x.ObrazProgramId, Name = x.ObrazProgramName, x.ObrazProgramCrypt }).Distinct().ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.ObrazProgramCrypt + " " + x.Name)).ToList();

                ComboServ.FillCombo(cbObrazProgram, src, false, false);
            }
        }
        private void FillComboProfile()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.qEntry.Where(x => (StudyLevelId.HasValue ? x.StudyLevelId == StudyLevelId : true)
                    && (StudyFormId.HasValue ? x.StudyFormId == StudyFormId : true)
                    && (LicenseProgramId.HasValue ? x.LicenseProgramId == LicenseProgramId : true)
                    && (ObrazProgramId.HasValue ? x.ObrazProgramId == ObrazProgramId : true)
                    && (FacultyId.HasValue ? x.FacultyId == FacultyId : true))
                    .Select(x => new { Id = x.ProfileId, Name = x.ProfileName }).Distinct().ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();

                ComboServ.FillCombo(cbProfile, src, false, false);
            }
        }

        public int? StudyLevelId
        {
            get { return ComboServ.GetComboIdInt(cbStudyLevel); }
            set { ComboServ.SetComboId(cbStudyLevel, value); }
        }
        public int? FacultyId
        {
            get { return ComboServ.GetComboIdInt(cbFaculty); }
            set { ComboServ.SetComboId(cbFaculty, value); }
        }
        public int? StudyFormId
        {
            get { return ComboServ.GetComboIdInt(cbStudyForm); }
            set { ComboServ.SetComboId(cbStudyForm, value); }
        }
        public int? LicenseProgramId
        {
            get { return ComboServ.GetComboIdInt(cbLicenseProgram); }
            set { ComboServ.SetComboId(cbLicenseProgram, value); }
        }
        public int? ObrazProgramId
        {
            get { return ComboServ.GetComboIdInt(cbObrazProgram); }
            set { ComboServ.SetComboId(cbObrazProgram, value); }
        }
        public Guid? ProfileId
        {
            get
            {
                Guid gRet = Guid.Empty;
                if (string.IsNullOrEmpty(ComboServ.GetComboId(cbProfile)) || !Guid.TryParse(ComboServ.GetComboId(cbProfile), out gRet))
                    return null;
                return gRet;
            }
            set
            {
                if (value.HasValue)
                    ComboServ.SetComboId(cbProfile, value.ToString());
            }
        }
        public Guid? EntryId { get; set; }

        private void cbStudyLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboFaculty();
        }
        private void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboLicenseProgram();
        }
        private void cbLicenseProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboObrazProgram();
        }
        private void cbObrazProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboProfile();
        }
        private void cbProfile_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboStudyForm();
        }
        private void cbStudyForm_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboStudyBasis();
        }
        private void cbStudyBasis_SelectedIndexChanged(object sender, EventArgs e)
        {
            using (PriemEntities context = new PriemEntities())
            {
                var EntryList = context.qEntry.Where(x => x.StudyBasisId == 2 && x.StudyLevelId == StudyLevelId && x.StudyFormId == StudyFormId && x.FacultyId == FacultyId
                    && x.LicenseProgramId == LicenseProgramId && x.ObrazProgramId == ObrazProgramId && (ProfileId.HasValue ? x.ProfileId == ProfileId : true)).Select(x => x.Id);

                if (EntryList.Count() == 1)
                    EntryId = EntryList.First();
            }
        }
        private void chbIsReduced_CheckedChanged(object sender, EventArgs e)
        {
            FillComboFaculty();
        }
        private void chbIsSecond_CheckedChanged(object sender, EventArgs e)
        {
            FillComboFaculty();
        }
        private void chbIsParallel_CheckedChanged(object sender, EventArgs e)
        {
            FillComboFaculty();
        }
        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (OnOK != null && EntryId.HasValue)
            {
                OnOK(EntryId.Value);
                this.Close();
            }
        }
    }
}
