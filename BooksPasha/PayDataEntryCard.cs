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
    public partial class PayDataEntryCard : BookCard
    {
        bool _isOpen = false;
        public PayDataEntryCard(int? studyLevelId, int? studyFormId, int? facultyId, int? licenseProgramId)
        {
            InitializeComponent();
            _tableName = "ed.PayDataEntry";
            InitControls();
            if (studyLevelId.HasValue)
                StudyLevelId = studyLevelId;
            if (studyFormId.HasValue)
                StudyFormId = studyFormId;
            else
                FillComboFaculty();
            if (facultyId.HasValue)
                FacultyId = facultyId;
            else
                FillComboLicenseProgram();
            if (licenseProgramId.HasValue)
            {
                LicenseProgramId = licenseProgramId;
                FillComboProfile();
                cbProfile_SelectedIndexChanged(null, null);
            }
            else
            {
                FillComboObrazProgram();
                FillComboProfile();
                cbProfile_SelectedIndexChanged(null, null);
            }
        }
        public PayDataEntryCard(Guid id)
        {
            InitializeComponent();
            _tableName = "ed.PayDataEntry";
            _Id = id.ToString();
            InitControls();
            //GuidId = id;
        }

        protected override void ExtraInit()
        {
            base.ExtraInit();
            CardTitle = "Реквизиты для договора по направлению";
            FillComboStudyLevel();
            if (GuidId.HasValue)
            {
                FillCard();
            }
        }

        private void OpenOrSave()
        {
            if (!_isOpen)
            {
                btnSaveChange.Text = "Сохранить";
                _isOpen = true;
            }
            else
            {
                btnSaveChange.Text = "Изменить";
                _isOpen = false;
            }
        }

        #region Handlers
        private void cbStudyLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboStudyForm();
        }
        private void cbStudyForm_SelectedIndexChanged(object sender, EventArgs e)
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
            FillComboProrector();
        }
        private void cbProfile_SelectedIndexChanged(object sender, EventArgs e)
        {
            using (PriemEntities context = new PriemEntities())
            {
                var EntryList = context.qEntry.Where(x => x.StudyBasisId == 2 && x.StudyLevelId == StudyLevelId && x.StudyFormId == StudyFormId && x.FacultyId == FacultyId
                    && x.LicenseProgramId == LicenseProgramId && x.ObrazProgramId == ObrazProgramId && (ProfileId.HasValue ? x.ProfileId == ProfileId : true)).Select(x => x.Id);

                if (EntryList.Count() == 0)
                    tbEntryId.Text = "!НЕ НАЙДЕНО!";
                else if (EntryList.Count() > 1)
                    tbEntryId.Text = "!НАЙДЕНО " + EntryList.Count().ToString();
                else
                {
                    tbEntryId.Text = EntryList.First().ToString();
                    EntryId = EntryList.First();
                }
            }
        }
        #endregion

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
        private void FillComboProrector()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.Prorektor.Select(x => new { Id = x.Id, Name = x.Name }).Distinct().ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();

                ComboServ.FillCombo(cbProrector, src, false, false);
            }
        }

        #region Fields
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
        public int? ProrectorId
        {
            get { return ComboServ.GetComboIdInt(cbProrector); }
            set { ComboServ.SetComboId(cbProrector, value); }
        }
        public int? ProfileId
        {
            get 
            {
                return ComboServ.GetComboIdInt(cbProfile);
            }
            set 
            { 
                if (value.HasValue)
                    ComboServ.SetComboId(cbProfile, value.ToString());
            }
        }
        public Guid? EntryId { get; set; }
        public string Props
        {
            get { return tbProps.Text.Trim(); }
            set { tbProps.Text = (value ?? "").Replace("\n", "\r\n").Replace("\r\r\n", "\r\n"); }
        }
        #endregion

        protected override void FillCard()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var data = context.PayDataEntry.Where(x => x.EntryId == GuidId)
                    .Select(x => new 
                    {
                        x.Entry.StudyLevelId,
                        x.Entry.StudyFormId,
                        x.Entry.FacultyId,
                        x.Entry.LicenseProgramId,
                        x.Entry.ObrazProgramId,
                        x.Entry.ProfileId,
                        x.ProrektorId,
                        x.Props,
                        x.EntryId
                    }).FirstOrDefault();

                if (data != null)
                {
                    StudyLevelId = data.StudyLevelId;
                    StudyFormId = data.StudyFormId;
                    FacultyId = data.FacultyId;
                    LicenseProgramId = data.LicenseProgramId;
                    ObrazProgramId = data.ObrazProgramId;
                    ProfileId = data.ProfileId;
                    ProrectorId = data.ProrektorId;
                    Props = data.Props;
                    EntryId = data.EntryId;
                }
            }
        }

        protected override void InsertRec(PriemEntities context, System.Data.Objects.ObjectParameter idParam)
        {
            context.PayDataEntry_Insert(EntryId, "Санкт-Петербургский государственный университет", 
                "199034, Санкт-Петербург, Университетская наб., д.7/9", 
                "7801002274",
                "", "", ProrectorId, "", "", new DateTime(2013, 9, 1), new DateTime(2017, 8, 31), Props);
            idParam.Value = EntryId;
        }
        protected override void UpdateRec(PriemEntities context, Guid id)
        {
            context.PayDataEntry_Update(EntryId, "Санкт-Петербургский государственный университет", 
                "199034, Санкт-Петербург, Университетская наб., д.7/9", 
                "7801002274",
                "", "", ProrectorId, "", "", new DateTime(2013, 9, 1), new DateTime(2017, 8, 31), Props);
        }
    }
}
