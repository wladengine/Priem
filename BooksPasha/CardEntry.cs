using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.Objects;

using EducServLib;

namespace Priem
{
    public partial class CardEntry : BookCard
    {
        #region Fields
        public int StudyFormId
        {
            get { return ComboServ.GetComboIdInt(cbStudyForm).Value; }
            set { ComboServ.SetComboId(cbStudyForm, value); }
        }
        public int StudyBasisId
        {
            get { return ComboServ.GetComboIdInt(cbStudyBasis).Value; }
            set { ComboServ.SetComboId(cbStudyBasis, value); }
        }
        public int StudyLevelId
        {
            get { return ComboServ.GetComboIdInt(cbStudyLevel).Value; }
            set { ComboServ.SetComboId(cbStudyLevel, value); }
        }
        public int? FacultyId
        {
            get { return ComboServ.GetComboIdInt(cbFaculty); }
            set { ComboServ.SetComboId(cbFaculty, value); }
        }
        public bool IsSecond
        {
            get { return chbIsSecond.Checked; }
            set { chbIsSecond.Checked = value; }
        }
        public bool IsParallel
        {
            get { return chbIsParallel.Checked; }
            set { chbIsParallel.Checked = value; }
        }
        public bool IsReduced
        {
            get { return chbIsReduced.Checked; }
            set { chbIsReduced.Checked = value; }
        }
        public int? AggregateGroupId
        {
            get { return ComboServ.GetComboIdInt(cbAggregateGroup); }
            set { ComboServ.SetComboId(cbAggregateGroup, value); }
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
        public int? ProfileId
        {
            get { return ComboServ.GetComboIdInt(cbProfile); }
            set { ComboServ.SetComboId(cbProfile, value); }
        }
        public int? ComissionId
        {
            get { return ComboServ.GetComboIdInt(cbComission); }
            set { ComboServ.SetComboId(cbComission, value); }
        }

        public DateTime DateOfStart
        {
            get { return dtpDateOfStart.Value; }
            set { dtpDateOfStart.Value = value; }
        }
        public DateTime DateOfClose
        {
            get { return dtpDateOfClose.Value; }
            set { dtpDateOfClose.Value = value; }
        }

        public DateTime? DateOfStart_Foreign
        {
            get 
            {
                if (dtpDateOfStart_Foreign.Checked)
                    return dtpDateOfStart_Foreign.Value;
                else
                    return null;
            }
            set 
            { 
                if (value.HasValue)
                    dtpDateOfStart_Foreign.Value = value.Value;

                dtpDateOfStart_Foreign.Checked = value.HasValue;
            }
        }
        public DateTime? DateOfClose_Foreign
        {
            get 
            {
                if (dtpDateOfClose_Foreign.Checked)
                    return dtpDateOfClose_Foreign.Value;
                else
                    return null;
            }
            set 
            {
                if (value.HasValue)
                    dtpDateOfClose_Foreign.Value = value.Value;

                dtpDateOfClose_Foreign.Checked = value.HasValue;
            }
        }

        public DateTime? DateOfStart_GosLine
        {
            get 
            {
                if (dtpDateOfStart_GosLine.Checked)
                    return dtpDateOfStart_GosLine.Value;
                else
                    return null;
            }
            set 
            { 
                if (value.HasValue)
                    dtpDateOfStart_GosLine.Value = value.Value;

                dtpDateOfStart_GosLine.Checked = value.HasValue;
            }
        }
        public DateTime? DateOfClose_GosLine
        {
            get 
            {
                if (dtpDateOfClose_GosLine.Checked)
                    return dtpDateOfClose_GosLine.Value;
                else
                    return null;
            }
            set 
            { 
                if (value.HasValue)
                    dtpDateOfClose_GosLine.Value = value.Value;

                dtpDateOfClose_GosLine.Checked = value.HasValue;
            }
        }

        public int KCP
        {
            get
            {
                int j;
                int.TryParse(tbKC.Text.Trim(), out j);
                return j;
            }
            set 
            {
                tbKC.Text = value.ToString();
            }
        }
        public int KCPCrimea
        {
            get
            {
                int j;
                int.TryParse(tbKCPCrimea.Text.Trim(), out j);
                return j;
            }
            set
            {
                tbKCPCrimea.Text = value.ToString();
            }
        }
        public int? KCPCel
        {
            get
            {
                int? iRet;
                int j;
                if (int.TryParse(tbKCPCel.Text.Trim(), out j))
                    iRet = j;
                else
                    iRet = null;

                return iRet;
            }
            set
            {
                if (value.HasValue)
                    tbKC.Text = value.ToString();
                else
                    tbKC.Text = "";
            }
        }
        public int? KCPQuota
        {
            get
            {
                int? iRet;
                int j;
                if (int.TryParse(tbKCPQuota.Text.Trim(), out j))
                    iRet = j;
                else
                    iRet = null;

                return iRet;
            }
            set
            {
                if (value.HasValue)
                    tbKCPQuota.Text = value.ToString();
                else
                    tbKCPQuota.Text = "";
            }
        }
        #endregion

        public CardEntry(string id)
            : base(id)
        {
            InitializeComponent();            
            InitControls();            
        }

        public CardEntry(int iStudyLevelId)
            : base()
        {
            InitializeComponent();
            InitControls();
            StudyLevelId = iStudyLevelId;
            UpdateAfterStudyLevel();
            UpdateAfterAggregateGroup();
            UpdateAfterLicenseProgram();
            UpdateAfterObrazProgram();
        }

        protected override void ExtraInit()
        {
            base.ExtraInit();
            _title = "Прием";
            _tableName = "ed.[Entry]";
            InitCombos();

            btnSaveAsNew.Visible = true;

            //DoSmth();
        }

        private void InitCombos()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.StudyForm.Select(x => new { x.Id, x.Name }).ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbStudyForm, src, false, false);
                src = context.StudyBasis.Select(x => new { x.Id, x.Name }).ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbStudyBasis, src, false, false);

                src = context.SP_Profile.Select(x => new { x.Id, x.Name }).ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbProfile, src, false, false);
                
                //src = context.SP_AggregateGroup.Select(x => new { x.Id, x.Name }).ToList()
                //    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                //ComboServ.FillCombo(cbAggregateGroup, src, false, false);

                src = context.StudyLevel.Select(x => new { x.Id, x.Name }).ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbStudyLevel, src, false, false);

                src = context.Comission.Select(x => new { x.Id, x.Name }).ToList()
                            .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbComission, src, true, false);

                UpdateAfterStudyLevel();
                UpdateAfterAggregateGroup();
                UpdateAfterLicenseProgram();
                UpdateAfterObrazProgram();
            }
        }

        //protected override void InitHandlers()
        //{
        //    cbStudyLevel.SelectedIndexChanged += cbStudyLevel_SelectedIndexChanged;
        //    cbLicenseProgram.SelectedIndexChanged += cbLicenseProgram_SelectedIndexChanged;
        //    cbObrazProgram.SelectedIndexChanged += cbObrazProgram_SelectedIndexChanged;
        //    cbProfile.SelectedIndexChanged += cbProfile_SelectedIndexChanged;
        //}
        //protected override void NullHandlers()
        //{
        //    cbStudyLevel.SelectedIndexChanged -= cbStudyLevel_SelectedIndexChanged;
        //    cbLicenseProgram.SelectedIndexChanged -= cbLicenseProgram_SelectedIndexChanged;
        //    cbObrazProgram.SelectedIndexChanged -= cbObrazProgram_SelectedIndexChanged;
        //    cbProfile.SelectedIndexChanged -= cbProfile_SelectedIndexChanged;
        //}

        private void UpdateAfterStudyLevel()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.SP_LicenseProgram.Where(x => x.StudyLevelId == StudyLevelId)
                    .Select(x => new { Id = x.AggregateGroupId, Name = x.SP_AggregateGroup.Name }).Distinct()
                    .ToList().Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbAggregateGroup, src, (src.Count == 0), false);
            }
        }
        private void UpdateAfterAggregateGroup()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.SP_LicenseProgram.Where(x => x.StudyLevelId == StudyLevelId && (AggregateGroupId.HasValue ? x.AggregateGroupId == AggregateGroupId : true))
                    .OrderBy(x => x.Code).Select(x => new { x.Id, x.Code, x.Name }).ToList()
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), "(" + x.Code + ") " + x.Name)).ToList();
                ComboServ.FillCombo(cbLicenseProgram, src, false, false);
            }
        }
        private void UpdateAfterLicenseProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.SP_ObrazProgram.Where(x => x.LicenseProgramId == LicenseProgramId).Select(x => new { x.Id, Acr1 = x.SP_LicenseProgram.StudyLevel.Acronym, x.Number, x.Name })
                    .ToList().OrderBy(x => x.Acr1).ThenBy(x => x.Number)
                    .Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Acr1 + "." + x.Number + "." + MainClass.sPriemYear + " " + x.Name)).ToList();
                ComboServ.FillCombo(cbObrazProgram, src, false, false);
            }
        }
        private void UpdateAfterObrazProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.SP_ObrazProgram.Where(x => x.LicenseProgramId == LicenseProgramId).Select(x => new { x.FacultyId, x.SP_Faculty.Name }).Distinct().ToList()
                    .Select(x => new KeyValuePair<string, string>(x.FacultyId.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbFaculty, src, false, false);
            }
        }

        protected override void SetIsOpen()
        {
            return;
        }
        protected override void DeleteIsOpen()
        {
            return;
        }
        protected override bool GetIsOpen()
        {
            return false;
        }     

        protected override void FillCard()
        {
            if (_Id == null)
            {
                return;
            }
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                   Entry ent = (from ex in context.Entry
                                 where ex.Id == GuidId
                                       select ex).FirstOrDefault();

                    StudyLevelId = ent.StudyLevelId;
                    UpdateAfterStudyLevel();
                    //tbStudyLevel.Text = ent.StudyLevelName;
                    AggregateGroupId = ent.SP_LicenseProgram.AggregateGroupId;
                    UpdateAfterAggregateGroup();
                    LicenseProgramId = ent.LicenseProgramId;
                    UpdateAfterLicenseProgram();
                    ObrazProgramId = ent.ObrazProgramId;
                    UpdateAfterObrazProgram();
                    ProfileId = ent.ProfileId;
                    FacultyId = ent.FacultyId;
                    StudyFormId = ent.StudyFormId;
                    StudyBasisId = ent.StudyBasisId;
                    tbStudyPlan.Text = ent.StudyPlanNumber;
                    tbKC.Text = ent.KCP.ToString();
                    IsSecond = ent.IsSecond;
                    IsReduced = ent.IsReduced;
                    IsParallel = ent.IsParallel;
                    tbKCPCel.Text = ent.KCPCel.ToString();
                    tbKCPCrimea.Text = ent.KCPCrimea.ToString();
                    KCPQuota = ent.KCPQuota;

                    ComissionId = ent.CommissionId;

                    DateOfStart = ent.DateOfStart;
                    DateOfClose = ent.DateOfClose;

                    DateOfStart_Foreign = ent.DateOfStart_Foreign;
                    DateOfClose_Foreign = ent.DateOfClose_Foreign;

                    DateOfStart_GosLine = ent.DateOfStart_GosLine;
                    DateOfClose_GosLine = ent.DateOfClose_GosLine;

                    UpdateExams();
                    UpdateInnerEntryInEntry();
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка при заполнении формы " + exc.Message);
            }
        }        

        protected override void SetReadOnlyFieldsAfterFill()
        {
            base.SetReadOnlyFieldsAfterFill();
            //if(_Id == null)
            //    btnSaveChange.Enabled = false;
            if (!MainClass.IsEntryChanger())
                btnSaveChange.Enabled = false;

            if (MainClass.IsEntryChanger())
            {
                WinFormsServ.SetSubControlsEnabled(tabPage4, true);
            }
        }

        protected override void SetAllFieldsEnabled()
        {
            base.SetAllFieldsEnabled();

            foreach (Control control in tcCard.Controls)
            {
                control.Enabled = true;
                foreach (Control crl in control.Controls)
                    crl.Enabled = true;
            }

            WinFormsServ.SetSubControlsEnabled(gbEntry, GetIsCanChangeEntry());
            tbKCPCel.Enabled = true;
            tbKC.Enabled = true;
            tbKCPCrimea.Enabled = true;
            tbKCPQuota.Enabled = true;
            cbComission.Enabled = true;
            cbFaculty.Enabled = true;
            tbStudyPlan.Enabled = true;
        }

        private bool GetIsCanChangeEntry()
        {
            if (MainClass.IsPasha() || MainClass.IsOwner())
                return true;
            if (string.IsNullOrEmpty(_Id) && ((MainClass.IsPasha() || MainClass.IsOwner() || MainClass.IsEntryChanger())))
                return true;
            using (PriemEntities context = new PriemEntities())
            {
                return (context.Abiturient.Where(x => x.EntryId == GuidId).Count() == 0);
            }
        }

        protected override void SetAllFieldsNotEnabled()
        {
            base.SetAllFieldsNotEnabled();
            tcCard.Enabled = true;

            foreach (Control control in tcCard.Controls)
            {
                foreach (Control crl in control.Controls)
                    crl.Enabled = false;
            }            
        }

        protected override void InsertRec(PriemEntities context, ObjectParameter idParam)
        {
            Guid Id = Guid.NewGuid();

            int _ProfileId = ComboServ.GetComboIdInt(cbProfile) ?? 0;


            var Entry = new Entry();
            Entry.Id = Id;
            Entry.FacultyId = FacultyId.Value;
            Entry.StudyLevelId = StudyLevelId;
            Entry.LicenseProgramId = LicenseProgramId.Value;
            Entry.ObrazProgramId = ObrazProgramId.Value;
            Entry.ProfileId = _ProfileId;
            Entry.StudyFormId = StudyFormId;
            Entry.StudyBasisId = StudyBasisId;
            if (!string.IsNullOrEmpty(tbStudyPlan.Text.Trim()))
                Entry.StudyPlanNumber = tbStudyPlan.Text.Trim();
            Entry.KCP = KCP;
            Entry.KCPCel = KCPCel;
            Entry.KCPQuota = KCPQuota;
            Entry.DateOfStart = DateOfStart;
            Entry.DateOfClose = DateOfClose;
            Entry.DateOfStart_Foreign = DateOfStart_Foreign;
            Entry.DateOfClose_Foreign = DateOfClose_Foreign;
            Entry.DateOfStart_GosLine = DateOfStart_GosLine;
            Entry.DateOfClose_GosLine = DateOfClose_GosLine;
            Entry.CommissionId = ComissionId;
            context.Entry.AddObject(Entry);
            context.SaveChanges();

            idParam.Value = Id;

            string query = @"INSERT INTO [_Entry] 
(Id, StudyLevelId, LicenseProgramId, ObrazProgramId, ProfileId, StudyFormId, StudyBasisId, FacultyId, SemesterId, CampaignYear,
DateOfStart, DateOfClose, DateOfStart_Foreign, DateOfClose_Foreign, DateOfStart_GosLine, DateOfClose_GosLine, ComissionId) VALUES
(@Id, @StudyLevelId, @LicenseProgramId, @ObrazProgramId, @ProfileId, @StudyFormId, @StudyBasisId, @FacultyId, @SemesterId, @CampaignYear,
@DateOfStart, @DateOfClose, @DateOfStart_Foreign, @DateOfClose_Foreign, @DateOfStart_GosLine, @DateOfClose_GosLine, @ComissionId)";
            SortedList<string, object> sl = new SortedList<string, object>();
            sl.Add("@Id", Id);

            sl.Add("@CampaignYear", MainClass.iPriemYear);
            sl.Add("@LicenseProgramId", LicenseProgramId);
            sl.Add("@ObrazProgramId", ObrazProgramId);
            sl.AddVal("@ProfileId", _ProfileId);
            sl.AddVal("@SemesterId", 1);

            sl.Add("@StudyLevelId", StudyLevelId);
            sl.Add("@StudyFormId", StudyFormId);
            sl.Add("@StudyBasisId", StudyBasisId);
            sl.AddVal("@FacultyId", FacultyId);
            sl.Add("@IsParallel", IsParallel);
            sl.Add("@IsReduced", IsReduced);
            sl.Add("@IsSecond", IsSecond);

            sl.AddVal("@DateOfStart_Foreign", DateOfStart_Foreign);
            sl.AddVal("@DateOfClose_Foreign", DateOfClose_Foreign);
            sl.AddVal("@DateOfStart_GosLine", DateOfStart_GosLine);
            sl.AddVal("@DateOfClose_GosLine", DateOfClose_GosLine);
            sl.AddVal("@DateOfStart", DateOfStart);
            sl.AddVal("@DateOfClose", DateOfClose);

            sl.AddVal("@ComissionId", ComissionId);

            MainClass.BdcOnlineReadWrite.ExecuteQuery(query, sl);
        }
        protected override void UpdateRec(PriemEntities context, Guid id)
        {
            context.Entry_UpdateCEl(GuidId, KCPCel);
            context.Entry_UpdateKC(GuidId, KCP, KCPCrimea, KCPQuota);
            context.Entry_Update(GuidId, StudyLevelId, StudyFormId, StudyBasisId, FacultyId, false, IsParallel, IsReduced, IsSecond, tbStudyPlan.Text.Trim(),
                DateOfStart_Foreign, DateOfClose_Foreign, DateOfStart_GosLine, DateOfClose_GosLine, DateOfStart, DateOfClose, ComissionId);

            try
            {
                string query = @"UPDATE [_Entry]
SET
    StudyLevelId=@StudyLevelId,
    StudyFormId=@StudyFormId,
    StudyBasisId=@StudyBasisId,
    FacultyId=@FacultyId,
    IsParallel=@IsParallel,
    IsReduced=@IsReduced,
    IsSecond=@IsSecond,
    DateOfStart_Foreign=@DateOfStart_Foreign,
	DateOfClose_Foreign=@DateOfClose_Foreign,
	DateOfStart_GosLine=@DateOfStart_GosLine,
	DateOfClose_GosLine=@DateOfClose_GosLine,
	DateOfStart=@DateOfStart,
	DateOfClose=@DateOfClose,
    CampaignYear=@CampaignYear,
    ComissionId=@ComissionId
WHERE Id=@Id";
                SortedList<string, object> sl = new SortedList<string, object>();
                sl.Add("@Id", GuidId.Value);

                sl.Add("@CampaignYear", MainClass.iPriemYear);
                sl.Add("@StudyLevelId", StudyLevelId);
                sl.Add("@StudyFormId", StudyFormId);
                sl.Add("@StudyBasisId", StudyBasisId);
                sl.AddVal("@FacultyId", FacultyId);
                sl.Add("@IsParallel", IsParallel);
                sl.Add("@IsReduced", IsReduced);
                sl.Add("@IsSecond", IsSecond);

                sl.AddVal("@DateOfStart_Foreign", DateOfStart_Foreign);
                sl.AddVal("@DateOfClose_Foreign", DateOfClose_Foreign);
                sl.AddVal("@DateOfStart_GosLine", DateOfStart_GosLine);
                sl.AddVal("@DateOfClose_GosLine", DateOfClose_GosLine);
                sl.AddVal("@DateOfStart", DateOfStart);
                sl.AddVal("@DateOfClose", DateOfClose);

                sl.AddVal("@ComissionId", ComissionId);

                MainClass.BdcOnlineReadWrite.ExecuteQuery(query, sl);
            }
            catch (Exception ex)
            {
                WinFormsServ.Error(ex);
            }
        }

        protected override void CloseCardAfterSave()
        {
            this.Close();
        }

        #region Exams

        public void UpdateExams()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var query = (from exEntry in context.extExamInEntry                             
                             where exEntry.EntryId == GuidId
                             orderby exEntry.IsProfil descending, exEntry.ExamName
                             select new 
                             { 
                                 exEntry.Id, 
                                 Name = exEntry.ExamName,
                                 IsProfil = exEntry.IsProfil ? "да" : "нет",
                                 exEntry.EgeMin
                             });

                dgvExams.DataSource = query;
                dgvExams.Columns["Id"].Visible = false;

                dgvExams.Columns["Name"].HeaderText = "Название"; 
                dgvExams.Columns["IsProfil"].HeaderText = "Профильный"; 
                dgvExams.Columns["EgeMin"].HeaderText = "Мин. ЕГЭ"; 
            }
        }

        private void OpenCardExam(Guid? entryId, string id, bool isForModified)
        {
            CardExamInEntry crd = new CardExamInEntry(entryId, id, isForModified);
            crd.ToUpdateList += new UpdateListHandler(UpdateExams);
            crd.Show();
        }


        private void btnOpenExam_Click(object sender, EventArgs e)
        {
            if (dgvExams.CurrentCell != null && dgvExams.CurrentCell.RowIndex > -1)
            {
                string itemId = dgvExams.Rows[dgvExams.CurrentCell.RowIndex].Cells["Id"].Value.ToString();
                if (!string.IsNullOrEmpty(itemId))
                    OpenCardExam(GuidId, itemId, _isModified);
            }
        }

        private void btnAddExam_Click(object sender, EventArgs e)
        {
            OpenCardExam(GuidId, null, true);
        }

        private void btnDeleteExam_Click(object sender, EventArgs e)
        {
            if (dgvExams.CurrentCell != null && dgvExams.CurrentCell.RowIndex > -1)
            {
                if (MessageBox.Show("Удалить запись?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    string itemId = dgvExams.CurrentRow.Cells["Id"].Value.ToString();
                    try
                    {
                        using (PriemEntities context = new PriemEntities())
                        {
                            int? id = int.Parse(itemId);
                            context.ExamInEntry_Delete(id);
                        }
                    }
                    catch (Exception ex)
                    {
                        WinFormsServ.Error("Каскадное удаление запрещено: " + ex.Message);
                    }

                    UpdateExams();
                }
            }
        }

        private void dgvExams_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            btnOpenExam_Click(null, null);
        }

        #endregion   

        #region InnerEntryInEntry
        private void UpdateInnerEntryInEntry()
        {
            if (string.IsNullOrEmpty(_Id))
                return;

            using (PriemEntities context = new PriemEntities())
            {
                var src = context.InnerEntryInEntry.Where(x => x.EntryId == GuidId)
                    .Select(x => new { x.Id, x.SP_ObrazProgram.Name, Profile = x.SP_Profile.Name, x.KCP }).ToArray();
                dgvInnerEntryInEntry.DataSource = Util.ConvertToDataTable(src);
                dgvInnerEntryInEntry.Columns["Id"].Visible = false;
                dgvInnerEntryInEntry.Columns["Name"].HeaderText = "Образовательная программа";
                dgvInnerEntryInEntry.Columns["Profile"].HeaderText = "Профиль";
                dgvInnerEntryInEntry.Columns["KCP"].HeaderText = "КЦП";
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (!GuidId.HasValue)
            {
                WinFormsServ.Error("Сохраните сперва карточку!");
                return;
            }

            if (LicenseProgramId.HasValue)
            {
                var crd = new CardInnerEntryInEntry(GuidId.Value, LicenseProgramId.Value);
                crd.ToUpdateList += UpdateInnerEntryInEntry;
                crd.Show();
            }
        }
        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dgvInnerEntryInEntry.SelectedCells.Count == 0)
                return;

            int rwInd = dgvInnerEntryInEntry.SelectedCells[0].RowIndex;
            Guid gId = (Guid)dgvInnerEntryInEntry["Id", rwInd].Value;

            using (PriemEntities context = new PriemEntities())
            {
                var OPIE = context.InnerEntryInEntry.Where(x => x.Id == gId).FirstOrDefault();
                context.InnerEntryInEntry.DeleteObject(OPIE);
                context.SaveChanges();

                string query = "DELETE FROM InnerEntryInEntry WHERE Id=@Id";
                SortedList<string, object> slParams = new SortedList<string, object>();
                slParams.Add("@Id", gId);
                MainClass.BdcOnlineReadWrite.ExecuteQuery(query, slParams);
            }

            UpdateInnerEntryInEntry();
        }
        private void dgvObrazProgramInEntry_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            Guid gId = (Guid)dgvInnerEntryInEntry["Id", e.RowIndex].Value;
            var crd = new CardInnerEntryInEntry(gId);
            crd.ToUpdateList += UpdateInnerEntryInEntry;
            crd.Show();
        }

        #endregion

        protected override bool CheckFields()
        {
            bool bRet = true;
            if (!LicenseProgramId.HasValue)
            {
                epError.SetError(cbLicenseProgram, "не указано направление");
                bRet = false;
            }
            if (!ObrazProgramId.HasValue)
            {
                epError.SetError(cbObrazProgram, "не указана обр. программа");
                bRet = false;
            }

            return bRet;
        }

        private void cbStudyLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateAfterStudyLevel();
        }
        private void cbLicenseProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateAfterLicenseProgram();
        }
        private void cbObrazProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateAfterObrazProgram();
        }
        private void cbProfile_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void cbAggregateGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateAfterAggregateGroup();
        }

        //private void DoSmth()
        //{
        //    DevelopmentStudentService.SQL_StudentSoapClient client = new DevelopmentStudentService.SQL_StudentSoapClient();
        //    var students = client.GetStudents("FIO", new System.Data.SqlTypes.SqlDateTime(DateTime.Now), 111);
        //}
    }
}
