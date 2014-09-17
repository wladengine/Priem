﻿using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Transactions;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.Objects;

using EducServLib;
using BDClassLib;
using WordOut;

using RtfWriter;

namespace Priem
{
    public partial class RatingList : BookList
    {
        string _queryFrom;
        string _queryBody;
        string _queryOlymps;
        string _queryFWJoin;
        string _queryOrange;

        bool bMagAddNabor1Enabled = false;
        DateTime dtMagAddNabor1 = new DateTime(2014, 8, 15);

        bool b1kursAddNabor1Enabled = true;
        DateTime dt1kursAddNabor1 = new DateTime(2014, 8, 15);

        bool bFirstWaveEnabled = true;

        //constructor
        public RatingList(bool fromFixieren)
        {
            InitializeComponent();
            InitVariables();

            _queryBody = @"SELECT DISTINCT ed.qAbiturient.Id as Id, ed.qAbiturient.RegNum as Рег_Номер, 
                    ed.extPerson.PersonNum as 'Ид. номер', ed.extPerson.FIO as ФИО, 
                    ed.extAbitMarksSum.TotalSum as 'Сумма баллов', ed.extAbitMarksSum.TotalCount as 'Кол-во оценок', 
                    case when EXISTS (SELECT * FROM ed.Abiturient AB WHERE AB.HasOriginals>0 AND AB.PersonId = qAbiturient.PersonId AND AB.BackDoc = 0) then 'Да' else 'Нет' end as 'Подлинники документов', 
                    ed.qAbiturient.Coefficient as 'Рейтинговый коэффициент', 
                    ed.Competition.Name as Конкурс, ed.hlpAbiturientProf.Prof AS 'Проф. экзамен', 
                    ed.hlpAbiturientProfAdd.ProfAdd AS 'Доп. экзамен',
                    extAbitMarkByRating.Value1 AS [Экзамен 1], extAbitMarkByRating.Value2 AS [Экзамен 2], extAbitMarkByRating.Value3 AS [Экзамен 3],
                    CASE WHEN EXISTS(SELECT Id FROM ed.hlpProfOlympiads AS Olympiads WHERE OlympValueId = 6 AND AbiturientId = ed.qAbiturient.Id) then 1 else CASE WHEN EXISTS(SELECT Id FROM ed.hlpProfOlympiads AS Olympiads WHERE OlympValueId = 5 AND AbiturientId = ed.qAbiturient.Id) then 2 else CASE WHEN EXISTS(SELECT Id FROM ed.hlpProfOlympiads AS Olympiads WHERE OlympValueId = 7 AND AbiturientId = ed.qAbiturient.Id) then 3 else 4 end end end as olymp,
                    CASE WHEN  ed.extPerson.AttestatSeries IN ('ЗА','ЗБ','ЗВ','АЗ') then 1 else CASE WHEN  ed.extPerson.AttestatSeries IN ('СА','СБ','СВ') then 2 else 3 end end as attestat,
                    (CASE WHEN ed.extPerson.IsExcellent=1 THEN 5 ELSE ed.extPerson.SchoolAVG END) as attAvg, 
                    CASE WHEN (CompetitionId=1  OR CompetitionId=8) then 1 else case when (CompetitionId=2 OR CompetitionId=7) AND ed.extPerson.Privileges>0 then 2 else (case when CompetitionId=6 then 3 else 4 end) end end as comp, 
                    CASE WHEN (CompetitionId=1 OR CompetitionId=8) then ed.qAbiturient.Coefficient else 10000 end as noexamssort, 
                    CASE WHEN (CompetitionId=5 OR CompetitionId=9) then 1 else 0 end as preimsort,
                    case when ed.extPerson.IsExcellent>0 then 'Да' else 'Нет' end as 'Медалист', 
                    ed.extPerson.AttestatSeries as 'Серия аттестата', 
                    ed.extPerson.DiplomSeries as 'Серия диплома', 
                    ed.extPerson.SchoolAVG as 'Средний балл', 
                    ed.extPerson.Email + ', '+ ed.extPerson.Phone + ', ' +ed.extPerson.Mobiles AS 'Контакты',
                    hlpAbiturientProf.Prof AS ProfSort, hlpAbiturientProfAdd.ProfAdd
                    /* (CASE WHEN hlpEntryWithAddExams.EntryId IS NULL THEN hlpAbiturientProf.Prof ELSE hlpAbiturientProfAdd.ProfAdd END) AS DopOrProfSort */";
 
            _queryOlymps = MainClass.studyLevelGroupId == 1 ? @", (SELECT TOP(1) ed.extOlympiads.OlympValueAcr + '-' + ed.extOlympiads.OlympName FROM ed.extOlympiads 
                           WHERE ed.extOlympiads.AbiturientId = ed.qAbiturient.Id AND ed.extOlympiads.OlympTypeId = 3 order by ed.extOlympiads.sortOrder) as 'Олимпиада' " : "";
             
            _queryFrom = @" FROM ed.qAbiturient 
                    INNER JOIN ed.extPerson ON ed.extPerson.Id = ed.qAbiturient.PersonId
                    INNER JOIN ed.Competition ON ed.Competition.Id = ed.qAbiturient.CompetitionId 
                    INNER JOIN ed.extEnableProtocol ON ed.extEnableProtocol.AbiturientId = ed.qAbiturient.Id 
                    LEFT JOIN ed.hlpEntryWithAddExams ON hlpEntryWithAddExams.EntryId = qAbiturient.EntryId
                    LEFT JOIN ed.hlpAbiturientProfAdd ON ed.hlpAbiturientProfAdd.Id = ed.qAbiturient.Id 
                    LEFT JOIN ed.hlpAbiturientProf ON ed.hlpAbiturientProf.Id = ed.qAbiturient.Id 
                    LEFT JOIN ed.extAbitMarksSum ON ed.qAbiturient.Id = ed.extAbitMarksSum.Id
                    LEFT JOIN ed.extAbitMarkByRating ON ed.qAbiturient.Id = ed.extAbitMarkByRating.Id
                    LEFT JOIN ed.hlpMinMarkAbiturient ON hlpMinMarkAbiturient.Id = qAbiturient.Id
                    LEFT JOIN ed.qAbiturientForeignApplicationsOnly qFor ON qAbiturient.Id = qFor.Id
                    LEFT JOIN ed.[_FirstWaveBackUp] FW ON FW.AbiturientId = qAbiturient.Id";

            if (MainClass.dbType == PriemType.PriemMag)
                _queryFrom += " LEFT JOIN ed.hlpMinMarkMag ON hlpMinMarkMag.AbiturientId = qAbiturient.Id";

            _queryFWJoin = "";                       
            
            Dgv = dgvAbits;
            _title = "Рейтинговый список";

            chbFix.Checked = fromFixieren;

            InitControls();            

            btnAdd.Visible = btnCard.Visible = btnRemove.Visible = false;
        }

        private void InitVariables()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var dicSettings = context.C_AppSettings.Select(x => new { x.ParamKey, x.ParamValue }).ToList().ToDictionary(x => x.ParamKey, y => y.ParamValue);
                string tmp = dicSettings.ContainsKey("bMagAddNabor1Enabled") ? dicSettings["bMagAddNabor1Enabled"] : "False";
                bMagAddNabor1Enabled = bool.Parse(tmp);

                tmp = dicSettings.ContainsKey("b1kursAddNabor1Enabled") ? dicSettings["b1kursAddNabor1Enabled"] : "False";
                b1kursAddNabor1Enabled = bool.Parse(tmp);

                tmp = dicSettings.ContainsKey("bFirstWaveEnabled") ? dicSettings["bFirstWaveEnabled"] : "False";
                bFirstWaveEnabled = bool.Parse(tmp);

                tmp = dicSettings.ContainsKey("dtMagAddNabor1") ? dicSettings["dtMagAddNabor1"] : new DateTime(DateTime.Now.Year, 8, 15).ToShortDateString();
                dtMagAddNabor1 = DateTime.Parse(tmp);

                tmp = dicSettings.ContainsKey("dt1kursAddNabor1") ? dicSettings["dt1kursAddNabor1"] : new DateTime(DateTime.Now.Year, 8, 15).ToShortDateString();
                dt1kursAddNabor1 = DateTime.Parse(tmp);

                if (MainClass.IsOwner() || MainClass.IsPasha())
                {
                    context.SetApplicationValue("bMagAddNabor1Enabled", bMagAddNabor1Enabled.ToString());
                    context.SetApplicationValue("b1kursAddNabor1Enabled", b1kursAddNabor1Enabled.ToString());
                    context.SetApplicationValue("bFirstWaveEnabled", bFirstWaveEnabled.ToString());
                    //dates
                    context.SetApplicationValue("dtMagAddNabor1", dtMagAddNabor1.ToShortDateString());
                    context.SetApplicationValue("dt1kursAddNabor1", dt1kursAddNabor1.ToShortDateString());
                }
            }
        }

        #region Init
        
        protected override void ExtraInit()
        {
            base.ExtraInit();

            btnFixieren.Visible = btnFixieren.Enabled = false;
            gbPasha.Visible = gbPasha.Enabled = false;
            chbFix.Visible = false;  

            if (MainClass.RightsFacMain() || MainClass.IsPasha())
                btnFixieren.Visible = btnFixieren.Enabled = true;

            if (MainClass.IsPasha())
            {
                gbPasha.Visible = gbPasha.Enabled = true;
                chbFix.Visible = true;  
            }

            if (!chbFix.Checked)
                gbPasha.Visible = gbPasha.Enabled = false;
            
            ComboServ.FillCombo(cbFaculty, HelpClass.GetComboListByTable("ed.qFaculty", "ORDER BY Acronym"), false, false);
            ComboServ.FillCombo(cbStudyBasis, HelpClass.GetComboListByTable("ed.StudyBasis", "ORDER BY Name"), false, false);

            cbStudyBasis.SelectedIndex = 0;
            FillStudyForm();
            FillLicenseProgram();
            FillObrazProgram();
            FillProfile();

            //если 
            //chbCel.Visible = false;

            if (MainClass.dbType == PriemType.PriemMag)
                chbWithOlymps.Visible = false;
        }

        public int? FacultyId
        {
            get { return ComboServ.GetComboIdInt(cbFaculty); }
            set { ComboServ.SetComboId(cbFaculty, value); }
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
                string prId = ComboServ.GetComboId(cbProfile);
                if (string.IsNullOrEmpty(prId))
                    return null;
                else
                    return new Guid(prId);
            }
            set
            {
                if (value == null)
                    ComboServ.SetComboId(cbProfile, (string)null);
                else
                    ComboServ.SetComboId(cbProfile, value.ToString());
            }
        }
        public int? StudyBasisId
        {
            get { return ComboServ.GetComboIdInt(cbStudyBasis); }
            set { ComboServ.SetComboId(cbStudyBasis, value); }
        }
        public int? StudyFormId
        {
            get { return ComboServ.GetComboIdInt(cbStudyForm); }
            set { ComboServ.SetComboId(cbStudyForm, value); }
        }
        public bool IsSecond
        {
            get { return chbIsSecond.Checked; }
            set { chbIsSecond.Checked = value; }
        }
        public bool IsReduced
        {
            get { return chbIsReduced.Checked; }
            set { chbIsReduced.Checked = value; }
        }
        public bool IsParallel
        {
            get { return chbIsParallel.Checked; }
            set { chbIsParallel.Checked = value; }
        }
        public bool IsCel
        {
            get { return chbCel.Checked; }
            set { chbCel.Checked = value; }
        }
        public bool IsCrimea
        {
            get { return chbIsCrimea.Checked; }
            set { chbIsCrimea.Checked = value; }
        }
        public bool IsQuota
        {
            get { return chbIsQuota.Checked; }
            set { chbIsQuota.Checked = value; }
        }
        public Guid? EntryId
        {
            get
            {
                try
                {
                    using (PriemEntities context = new PriemEntities())
                    {
                        Guid? entId = (from ent in MainClass.GetEntry(context)
                                       where ent.IsSecond == IsSecond && ent.IsParallel == IsParallel && ent.IsReduced == IsReduced
                                       && ent.LicenseProgramId == LicenseProgramId
                                       && ent.ObrazProgramId == ObrazProgramId
                                       && (ProfileId == null ? ent.ProfileId == null : ent.ProfileId == ProfileId)
                                       && ent.StudyFormId == StudyFormId
                                       && ent.StudyBasisId == StudyBasisId
                                       select ent.Id).FirstOrDefault();
                        return entId;
                    }
                }
                catch
                {
                    return null;
                }
            }            
        }

        private void FillStudyForm()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = MainClass.GetEntry(context).Where(c => c.FacultyId == FacultyId).Where(c=>c.StudyBasisId == StudyBasisId);
                
                ent = ent.Where(c => c.IsSecond == IsSecond && c.IsReduced == IsReduced && c.IsParallel == IsParallel);

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.StudyFormId.ToString(), u.StudyFormName)).Distinct().ToList();

                ComboServ.FillCombo(cbStudyForm, lst, false, false);                
            }
        }
        private void FillLicenseProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = MainClass.GetEntry(context).Where(c => c.FacultyId == FacultyId);

                ent = ent.Where(c => c.IsSecond == IsSecond && c.IsReduced == IsReduced && c.IsParallel == IsParallel);

                if (StudyBasisId != null)
                    ent = ent.Where(c => c.StudyBasisId == StudyBasisId);
                if (StudyFormId != null)
                    ent = ent.Where(c => c.StudyFormId == StudyFormId);

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.LicenseProgramId.ToString(), u.LicenseProgramName)).Distinct().ToList();

                ComboServ.FillCombo(cbLicenseProgram, lst, false, false);                
            }
        }
        private void FillObrazProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = MainClass.GetEntry(context).Where(c => c.FacultyId == FacultyId);

                ent = ent.Where(c => c.IsSecond == IsSecond && c.IsReduced == IsReduced && c.IsParallel == IsParallel);

                if (StudyBasisId != null)
                    ent = ent.Where(c => c.StudyBasisId == StudyBasisId);
                if (StudyFormId != null)
                    ent = ent.Where(c => c.StudyFormId == StudyFormId);
                if (LicenseProgramId != null)
                    ent = ent.Where(c => c.LicenseProgramId == LicenseProgramId);

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.ObrazProgramId.ToString(), u.ObrazProgramName + ' ' + u.ObrazProgramCrypt)).Distinct().ToList();

                ComboServ.FillCombo(cbObrazProgram, lst, false, false);
            }
        }
        private void FillProfile()
        {
            using (PriemEntities context = new PriemEntities())
            {
                if (ObrazProgramId == null)
                {
                    ComboServ.FillCombo(cbProfile, new List<KeyValuePair<string, string>>(), false, false);
                    cbProfile.Enabled = false;
                    return;
                }

                var ent = MainClass.GetEntry(context).Where(c => c.FacultyId == FacultyId).Where(c => c.ProfileId != null);

                ent = ent.Where(c => c.IsSecond == IsSecond && c.IsReduced == IsReduced && c.IsParallel == IsParallel);

                if (StudyBasisId != null)
                    ent = ent.Where(c => c.StudyBasisId == StudyBasisId);
                if (StudyFormId != null)
                    ent = ent.Where(c => c.StudyFormId == StudyFormId);
                if (LicenseProgramId != null)
                    ent = ent.Where(c => c.LicenseProgramId == LicenseProgramId);
                if (ObrazProgramId != null)
                    ent = ent.Where(c => c.ObrazProgramId == ObrazProgramId);

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.ProfileId.ToString(), u.ProfileName)).Distinct().ToList();

                if (lst.Count() > 0)
                {
                    ComboServ.FillCombo(cbProfile, lst, false, false);
                    cbProfile.Enabled = true;
                }
                else
                {
                    ComboServ.FillCombo(cbProfile, new List<KeyValuePair<string, string>>(), false, false);
                    cbProfile.Enabled = false;
                }              
            }
        }
        
        //инициализация обработчиков мегакомбов
        public override void InitHandlers()
        {
            cbFaculty.SelectedIndexChanged += new EventHandler(cbFaculty_SelectedIndexChanged);
            cbStudyForm.SelectedIndexChanged += new EventHandler(cbStudyForm_SelectedIndexChanged);
            cbStudyBasis.SelectedIndexChanged += new EventHandler(cbStudyBasis_SelectedIndexChanged);
            cbLicenseProgram.SelectedIndexChanged += new EventHandler(cbLicenseProgram_SelectedIndexChanged);
            cbObrazProgram.SelectedIndexChanged += new EventHandler(cbObrazProgram_SelectedIndexChanged);
            
            chbFix.CheckedChanged += new EventHandler(chbFix_CheckedChanged);
        }

        void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {            
            FillStudyForm();
            NullDataGrid();
        }
        void cbStudyBasis_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillStudyForm();
            NullDataGrid();
        }
        void cbStudyForm_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillLicenseProgram();
            NullDataGrid();
        }        
        void cbLicenseProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillObrazProgram();
            NullDataGrid();
        }
        void cbObrazProgram_SelectedIndexChanged(object sender, EventArgs e)
        {           
            FillProfile();
            NullDataGrid();
        }       
        private void chbFix_CheckedChanged(object sender, EventArgs e)
        {
            if (chbFix.Checked)
                gbPasha.Visible = gbPasha.Enabled = true;
            else
                gbPasha.Visible = gbPasha.Enabled = false;

            UpdateDataGrid();
        }
 
        #endregion
        
        protected override void OpenCard(string id, BaseFormsLib.BaseFormEx formOwner, int? index)
        {
            MainClass.OpenCardAbit(id, this, dgvAbits.CurrentRow.Index);
        }

        int GetPlanValueAndCheckLock()
        {
            using (PriemEntities context = new PriemEntities())
            {
                int plan = 0, planCel = 0, planCrimea = 0, planQuota = 0, entered = 0, enteredCel = 0, enteredCrimea = 0, enteredQuota = 0;               

                qEntry entry = (from ent in MainClass.GetEntry(context)
                       where ent.IsReduced == IsReduced && ent.IsParallel == IsParallel && ent.IsSecond == IsSecond 
                       && ent.FacultyId == FacultyId && ent.LicenseProgramId == LicenseProgramId
                       && ent.ObrazProgramId == ObrazProgramId
                       && (ProfileId == null ? ent.ProfileId == null : ent.ProfileId == ProfileId)
                       && ent.StudyFormId == StudyFormId
                       && ent.StudyBasisId == StudyBasisId
                       select ent).FirstOrDefault();

                if (entry == null)
                    return 0;

                plan = entry.KCP ?? 0;
                planCel = entry.KCPCel ?? 0;
                planCrimea = entry.KCPCrimea ?? 0;
                planQuota = entry.KCPQuota ?? 0;

                Guid? entryId = entry.Id;

                entered = (from ab in context.qAbitAll
                           join ev in context.extEntryView
                           on ab.Id equals ev.AbiturientId
                           where ab.CompetitionId != 6 && ab.CompetitionId != 11 && ab.CompetitionId != 12 && ab.CompetitionId != 2 && ab.CompetitionId != 7 && ab.EntryId == entryId
                           select ab).Count();

                enteredCrimea = (from ab in context.qAbitAll
                           join ev in context.extEntryView
                           on ab.Id equals ev.AbiturientId
                           where (ab.CompetitionId == 11 || ab.CompetitionId == 12) && ab.EntryId == entryId
                           select ab).Count();

                enteredQuota = (from ab in context.qAbitAll
                                 join ev in context.extEntryView
                                 on ab.Id equals ev.AbiturientId
                                 where (ab.CompetitionId == 2 || ab.CompetitionId == 7) && ab.EntryId == entryId
                                 select ab).Count();

                enteredCel = (from ab in context.qAbitAll
                              join ev in context.extEntryView
                              on ab.Id equals ev.AbiturientId
                              where ab.CompetitionId == 6 && ab.EntryId == entryId
                              select ab).Count();
               
                CheckLockAndPasha(context);

                if (IsCel)
                    return planCel - enteredCel;
                else if (IsCrimea)
                    return planCrimea - enteredCrimea;
                else if (IsQuota)
                    return planQuota - enteredQuota;
                else
                    return plan - planCel - entered - enteredQuota;
            }
        }

        private void CheckLockAndPasha(PriemEntities context)
        {
            //лочит кнопку 
            FixierenView fixView = (from fv in context.FixierenView
                           where fv.StudyLevelGroupId == MainClass.studyLevelGroupId && fv.IsReduced == IsReduced && fv.IsParallel == IsParallel && fv.IsSecond == IsSecond
                           && fv.FacultyId == FacultyId && fv.LicenseProgramId == LicenseProgramId
                           && fv.ObrazProgramId == ObrazProgramId
                           && (ProfileId == null ? fv.ProfileId == null : fv.ProfileId == ProfileId)
                           && fv.StudyFormId == StudyFormId
                           && fv.StudyBasisId == StudyBasisId
                           && fv.IsCel == IsCel
                           && fv.IsCrimea == IsCrimea
                           select fv).FirstOrDefault();
            
            string DocNum = string.Empty;
            bool? locked = false;

            if (fixView != null)
            {
                DocNum = fixView.DocNum.ToString(); ;
                locked = fixView.Locked;
            }

            lblNumber.Text = DocNum.Length == 0 ? " -----" : DocNum;
            lblLocked.Text = locked.GetValueOrDefault(false) ? "ЗАЛОЧЕНА" : "НЕ залочена";

            return;
        }

        public void NullDataGrid()
        {
            if (dgvAbits.DataSource != null)
            {
                dgvAbits.DataSource = null;
                lblCount.Text = string.Empty;
            }
        }

        //обновление грида
        int plan = 0;
        public override void UpdateDataGrid()
        {
            try
            {                
                string sOrderBy = string.Empty;
                if (MainClass.dbType == PriemType.PriemMag)
                {
                    sOrderBy =
                        chbCel.Checked ?
                        " ORDER BY ed.qAbiturient.Coefficient, comp , noexamssort desc, ed.extAbitMarksSum.TotalSum desc, ed.extAbitMarksSum.TotalCount desc, ФИО" :
                        " ORDER BY comp , noexamssort, ed.extAbitMarksSum.TotalSum desc, 'Проф. экзамен' DESC, ed.qAbiturient.Coefficient DESC, ed.extAbitMarksSum.TotalCount desc, ФИО";
                }
                else
                {
                    sOrderBy =
                        chbCel.Checked ?
                        " ORDER BY ed.qAbiturient.Coefficient, comp, noexamssort desc, ed.extAbitMarksSum.TotalSum desc, ProfSort desc, ProfAdd desc, ed.extAbitMarksSum.TotalCount desc, ФИО"
                        :
                        " ORDER BY comp, noexamssort, ed.extAbitMarksSum.TotalSum desc, [Экзамен 1] desc, [Экзамен 2] desc, [Экзамен 3] desc, preimsort desc, ProfAdd desc, " +
                        "olymp, Медалист, attAvg desc, ed.qAbiturient.Coefficient, ed.extAbitMarksSum.TotalCount desc, ФИО"
                        ;
                }
                string totalQuery = null;

                
                plan = GetPlanValueAndCheckLock();

                if (chbFix.Checked)
                {
                    if (MainClass.dbType == PriemType.PriemMag)
                        _queryOrange = @", CASE WHEN EXISTS(SELECT PersonId FROM ed.hlpPersonsWithOriginals WHERE PersonId = ed.qAbiturient.PersonId AND EntryId <> ed.qAbiturient.EntryId) then 1 else 0 end as orange ";
                    else
                        _queryOrange = @", CASE WHEN EXISTS(SELECT ed.extEntryView.Id FROM ed.extEntryView INNER JOIN ed.Abiturient a ON ed.extEntryView.AbiturientId = a.Id WHERE a.PersonId = ed.qAbiturient.PersonId) then 1 else 0 end as orange ";

                    string queryFix = _queryBody + _queryOrange + 
                    @" FROM ed.qAbiturient                     
                    INNER JOIN ed.extPerson ON ed.extPerson.Id = ed.qAbiturient.PersonId                    
                    INNER JOIN ed.Competition ON ed.Competition.Id = ed.qAbiturient.CompetitionId 
                    INNER JOIN ed.Fixieren ON ed.Fixieren.AbiturientId=ed.qAbiturient.Id 
                    LEFT JOIN ed.hlpEntryWithAddExams ON hlpEntryWithAddExams.EntryId = qAbiturient.EntryId
                    LEFT JOIN ed.FixierenView ON ed.Fixieren.FixierenViewId=ed.FixierenView.Id 
                    LEFT JOIN ed.hlpAbiturientProfAdd ON ed.hlpAbiturientProfAdd.Id = ed.qAbiturient.Id 
                    LEFT JOIN ed.hlpAbiturientProf ON ed.hlpAbiturientProf.Id = ed.qAbiturient.Id 
                    LEFT JOIN ed.extAbitMarksSum ON ed.qAbiturient.Id = ed.extAbitMarksSum.Id
                    LEFT JOIN ed.extAbitMarkByRating ON ed.qAbiturient.Id = ed.extAbitMarkByRating.Id
                    
";

                    string whereFix = string.Format(
@" WHERE ed.FixierenView.StudyLevelGroupId = {10} AND ed.FixierenView.StudyFormId={0} AND ed.FixierenView.StudyBasisId={1} AND ed.FixierenView.FacultyId={2} 
AND ed.FixierenView.LicenseProgramId={3} AND ed.FixierenView.ObrazProgramId={4} {5} AND ed.FixierenView.IsCel = {6}
AND ed.FixierenView.IsSecond = {7} AND ed.FixierenView.IsReduced = {8} AND ed.FixierenView.IsParallel = {9} AND ed.FixierenView.IsCrimea = {11} AND ed.FixierenView.IsQuota = {12}",
                        StudyFormId, StudyBasisId, FacultyId, LicenseProgramId, ObrazProgramId, ProfileId == null ? " AND ed.FixierenView.ProfileId IS NULL" : "AND ed.FixierenView.ProfileId='" + ProfileId + "'",
                        QueryServ.StringParseFromBool(IsCel), QueryServ.StringParseFromBool(IsSecond), QueryServ.StringParseFromBool(IsReduced), QueryServ.StringParseFromBool(IsParallel), 
                        MainClass.studyLevelGroupId, QueryServ.StringParseFromBool(IsCrimea), QueryServ.StringParseFromBool(IsQuota));
                    
                    //sOrderBy = " ORDER BY Fixieren.Number ";

                    totalQuery = queryFix + whereFix + sOrderBy;
                }
                else
                {
                    string sFilters = GetFilterString();
                    
                    //целевики?
                    //if (chbCel.Checked)
                    //    sFilters += " AND ed.qAbiturient.CompetitionId IN (6) ";
                    // в общем списке выводить всех 
                    //else
                    //    sFilters += " AND ed.qAbiturient.CompetitionId NOT IN (6) ";

                    //не забрали доки
                    sFilters += " AND (ed.qAbiturient.BackDoc=0) ";

                    sFilters += " AND ed.qAbiturient.Id NOT IN (select abiturientid from ed.extentryview) ";

                    //не иностранцы
                    sFilters += " AND qFor.Id IS NULL ";

                    //крым?
                    if (IsCrimea)
                        sFilters += " AND ed.qAbiturient.CompetitionId IN (11, 12) ";
                    else
                        sFilters += " AND ed.qAbiturient.CompetitionId NOT IN (11, 12) ";

                    //квотники?
                    if (IsQuota)
                        sFilters += " AND ed.qAbiturient.CompetitionId IN (2, 7) ";
                    //else
                    //    sFilters += " AND ed.qAbiturient.CompetitionId NOT IN (2, 7) ";

                    // кроме бэ преодолены мин планки 
                    if (MainClass.dbType == PriemType.PriemMag)
                        sFilters += " AND ((CompetitionId=1 OR CompetitionId=8 OR CompetitionId=12) OR hlpMinMarkMag.AbiturientId IS NULL)";
                    else
                        sFilters += " AND ((CompetitionId=1 OR CompetitionId=8 OR CompetitionId=12) OR hlpMinMarkAbiturient.Id IS NULL)";

                    string examsCnt = _bdc.GetStringValue(string.Format(" SELECT Count(Id) FROM ed.extExamInEntry WHERE EntryId='{0}' ", EntryId.ToString()));
                   
                    if (MainClass.dbType == PriemType.PriemMag)
                    { 
                        _queryOrange = @", CASE WHEN EXISTS(SELECT PersonId FROM ed.hlpPersonsWithOriginals WHERE PersonId = ed.qAbiturient.PersonId AND EntryId <> ed.qAbiturient.EntryId) then 1 else 0 end as orange ";
                        
                        // кроме бэ нужное кол-во оценок есть
                        sFilters += " AND ((CompetitionId=1  OR CompetitionId=8 OR CompetitionId=12) OR ed.extAbitMarksSum.TotalCount = " + examsCnt + " ) ";

                        if (bMagAddNabor1Enabled)
                            sFilters += " AND qAbiturient.DocInsertDate > '" + dtMagAddNabor1.ToShortDateString() + "' ";

                        totalQuery = _queryBody + _queryOrange + _queryFrom + sFilters + sOrderBy;
                    }
                    else
                    {
                        _queryOrange = @", CASE WHEN EXISTS(SELECT ed.extEntryView.Id FROM ed.extEntryView INNER JOIN ed.Abiturient a ON ed.extEntryView.AbiturientId = a.Id WHERE a.PersonId = ed.qAbiturient.PersonId) then 1 else 0 end as orange ";

                        if (bFirstWaveEnabled && MainClass.dbType == PriemType.Priem && StudyBasisId != 2)
                            sFilters += " AND FW.AbiturientId IS NOT NULL";

                        //эту хрень использовать только во второй волне - оно не будет работать, пока в _FirstWaveBackup или в _FirstWave не появятся люди
                        //ЗАКОММЕНТИТЬ К НОВОМУ ПРИЁМУ!!!
                        //if (StudyBasisId == 2)
                        //    sFilters += " AND qAbiturient.Id IN (SELECT AbiturientId FROM ed.[_FirstWaveBackup1-29072013])"; 
                        //else
                        //    sFilters += " AND qAbiturient.Id IN (SELECT AbiturientId FROM ed.[_FirstWave])";

                        if (MainClass.dbType == PriemType.Priem && b1kursAddNabor1Enabled)
                            sFilters += " AND qAbiturient.DocInsertDate > '" + dt1kursAddNabor1.ToShortDateString() + "' ";

                        //до зачисления льготников выводить их, а потом - убирать
                        //sFilters += " AND ed.qAbiturient.CompetitionId NOT IN (6, 1, 2, 7) ";                                        

                        // кроме бэ и тех, у кого нет сертификатов и оценок нужное кол-во оценок есть
                        sFilters += @" AND (CompetitionId IN (1, 8, 11, 12) 
                                        OR (ed.qAbiturient.PersonId NOT IN (SELECT PersonId FROM ed.EgeCertificate) 
                                           AND ed.qAbiturient.Id NOT IN (SELECT abiturientid from ed.Mark where IsFromEge = 1) /*and ed.extPerson.EgeInSPbgu = 0 */and ed.qAbiturient.IsSecond = 0 and ed.qAbiturient.IsReduced = 0 and ed.qAbiturient.IsParallel = 0) 
                                        OR ed.extAbitMarksSum.TotalCount = " + examsCnt + " ) ";
                        
                        //if (StudyBasisId == 2)
                        //    sFilters += " AND qAbiturient.Id NOT IN (SELECT AbiturientId FROM ed._FirstWaveGreen)";

                        totalQuery = _queryBody + (chbWithOlymps.Checked ? _queryOlymps : "") + _queryOrange + _queryFrom + sFilters + sOrderBy;
                    }
                }

                if (!dgvAbits.Columns.Contains("Number"))
                    dgvAbits.Columns.Add("Number", "№ п/п");

                HelpClass.FillDataGrid(dgvAbits, _bdc, totalQuery, "");

                dgvAbits.Columns["Id"].Visible = false;
                dgvAbits.Columns["comp"].Visible = false;
                dgvAbits.Columns["noexamssort"].Visible = false;
                dgvAbits.Columns["preimsort"].Visible = false;
                dgvAbits.Columns["olymp"].Visible = false;
                dgvAbits.Columns["attestat"].Visible = false;
                dgvAbits.Columns["attAvg"].Visible = false;
                dgvAbits.Columns["ProfSort"].Visible = false;
                dgvAbits.Columns["ProfAdd"].Visible = false;
                dgvAbits.Columns["orange"].Visible = false;

                if (MainClass.dbType == PriemType.PriemMag)
                {
                    dgvAbits.Columns["Серия аттестата"].Visible = false;
                    dgvAbits.Columns["Медалист"].HeaderText = "Красный диплом";
                }
                else
                    dgvAbits.Columns["Серия диплома"].Visible = false;
                
                foreach (DataGridViewColumn column in dgvAbits.Columns)
                {
                    column.SortMode = DataGridViewColumnSortMode.NotSortable;
                }

                lblCount.Text = dgvAbits.RowCount.ToString() + "             Cвободных мест: "+plan;
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при обновлении списка.");
            }
        }
      
        private string GetFilterString()
        {
            string s = " WHERE 1=1 ";
            s += " AND ed.qAbiturient.StudyLevelGroupId = " + MainClass.studyLevelGroupId;  
            
            //s += " AND ed.qAbiturient.DocDate>='20120813'"; 

            //обработали факультет
            if (FacultyId != null)
                s += " AND ed.qAbiturient.FacultyId = " + FacultyId;      
            
            //обработали форму обучения  
            if (StudyFormId != null)
                s += " AND ed.qAbiturient.StudyFormId = " + StudyFormId;

            //обработали основу обучения  
            if (StudyBasisId != null)
                s += " AND ed.qAbiturient.StudyBasisId = " + StudyBasisId;               

            //обработали Направление
            if (LicenseProgramId != null)
                s += " AND ed.qAbiturient.LicenseProgramId = " + LicenseProgramId;

            //обработали Образ программу
            if (ObrazProgramId != null)
                s += " AND ed.qAbiturient.ObrazProgramId = " + ObrazProgramId;

            //обработали профиль
            if (ProfileId != null)
                s += string.Format(" AND ed.qAbiturient.ProfileId = '{0}'", ProfileId);
            else
                s += " AND ed.qAbiturient.ProfileId IS NULL";


            s += " AND ed.qAbiturient.IsSecond = " + (IsSecond ? " 1 " : " 0 ");
            s += " AND ed.qAbiturient.IsReduced = " + (IsReduced ? " 1 " : " 0 ");
            s += " AND ed.qAbiturient.IsParallel = " + (IsParallel ? " 1 " : " 0 ");

            if (chbCel.Checked)
                s += " AND ed.qAbiturient.CompetitionId = 6 ";

            return s;
        }

        private void dgvAbits_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex == dgvAbits.Columns["Number"].Index)
            {
                e.Value = string.Format("{0}", e.RowIndex + 1);
            }

            if (e.RowIndex < plan)
            {
                if (e.ColumnIndex != dgvAbits.Columns["ФИО"].Index)//сперва подкрасим не-фио
                    dgvAbits[e.ColumnIndex, e.RowIndex].Style.BackColor = System.Drawing.Color.LightGreen;
                //потом докрашиваем не-оранжевые фио
                if (e.ColumnIndex == dgvAbits.Columns["ФИО"].Index && dgvAbits["orange", e.RowIndex].Value.ToString() != "1")
                    dgvAbits[e.ColumnIndex, e.RowIndex].Style.BackColor = System.Drawing.Color.LightGreen;
            }
            //и в последнюю очередь - оранжевых
            //это позволяет избежать рекурсивного вызова "перекраски" (сперва ячейка зелёная, а потом оранжевая)
            if (e.ColumnIndex == dgvAbits.Columns["ФИО"].Index && dgvAbits["orange", e.RowIndex].Value.ToString() == "1")
            {
                dgvAbits["ФИО", e.RowIndex].Style.BackColor = System.Drawing.Color.Orange;
            }            
        }

        private void tbNumber_TextChanged(object sender, EventArgs e)
        {
            WinFormsServ.Search(this.dgvAbits, "Рег_номер", tbNumber.Text);
        }

        private void tbFIO_TextChanged(object sender, EventArgs e)
        {
            WinFormsServ.Search(this.dgvAbits, "ФИО", tbFIO.Text);
        }

        private void btnFixieren_Click(object sender, EventArgs e)
        {
            Fixieren();
        }        

        private void Fixieren()
        {
            if (dgvAbits.DataSource == null || dgvAbits.Rows.Count == 0)
                return;

            using (PriemEntities context = new PriemEntities())
            {
                using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew))
                {
                    try
                    {
                        Guid? fixViewId = (from fv in context.FixierenView
                                           where fv.StudyLevelGroupId == MainClass.studyLevelGroupId && fv.IsReduced == IsReduced && fv.IsParallel == IsParallel && fv.IsSecond == IsSecond
                                           && fv.FacultyId == FacultyId && fv.LicenseProgramId == LicenseProgramId
                                           && fv.ObrazProgramId == ObrazProgramId
                                           && (ProfileId == null ? fv.ProfileId == null : fv.ProfileId == ProfileId)
                                           && fv.StudyFormId == StudyFormId
                                           && fv.StudyBasisId == StudyBasisId
                                           && fv.IsCel == IsCel
                                           && fv.IsCrimea == IsCrimea
                                           && fv.IsQuota == IsQuota
                                           select fv.Id).FirstOrDefault();

                        if (fixViewId != null)
                        {
                            bool? locked = (from fv in context.FixierenView
                                            where fv.Id == fixViewId
                                            select fv.Locked).FirstOrDefault();

                            if (locked.HasValue && locked.Value)
                            {
                                WinFormsServ.Error("Создание представления заблокировано, т.к. уже утверждена предыдущая версия");
                                return;
                            }

                            context.Fixieren_DeleteByFVId(fixViewId);
                            context.FixierenView_Delete(fixViewId);
                        }

                        int rand = new Random().Next(10000, 99999);

                        ObjectParameter fvId = new ObjectParameter("id", typeof(Guid));
                        context.FixierenView_Insert(MainClass.studyLevelGroupId, FacultyId, LicenseProgramId, ObrazProgramId, ProfileId, StudyBasisId, StudyFormId, IsSecond, IsReduced, IsParallel, IsCel, rand, false, IsCrimea, IsQuota, fvId);
                        Guid? viewId = (Guid?)fvId.Value;

                        int counter = 0;
                        foreach (DataGridViewRow row in dgvAbits.Rows)
                        {
                            counter++;
                            Guid? abId = new Guid(row.Cells["Id"].Value.ToString());
                            context.Fixieren_Insert(counter, abId, viewId);
                        }

                        transaction.Complete();                        
                    }
                    catch (Exception ex)
                    {
                        WinFormsServ.Error("Ошибка при сохранении списка");
                        return;
                    }                   
                }

                //ПЕЧАТЬ!
                PrintProtocol();
            }             
        }

        private void PrintProtocol()
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "ADOBE Pdf files|*.pdf";
            if (sfd.ShowDialog() == DialogResult.OK)
                Print.PrintRatingProtocol(StudyFormId, StudyBasisId, FacultyId, LicenseProgramId, ObrazProgramId, ProfileId, IsCel, IsCrimea, 
                    plan, sfd.FileName, IsSecond, IsReduced, IsParallel, IsQuota);
        }        

        private void btnWord_Click(object sender, EventArgs e)
        {
            ToWord();
        }

        private void ToWord()
        {
            int rowCount = dgvAbits.Rows.Count;
            if (rowCount == 0)
                return;
            try
            {
                float margin = (float)(20.0m * RtfConstants.MILLIMETERS_TO_POINTS);
                RtfDocument doc = new RtfDocument(PaperSize.A4, PaperOrientation.Landscape, Lcid.Russian, margin, margin, margin, margin);

                RtfTable table = doc.addTable(rowCount + 1, 14, (float)(276.1m * RtfConstants.MILLIMETERS_TO_POINTS));

                // Устанавливаем ширину столбцов таблицы (в миллиметрах)
                table.setColWidth(0, (float)(8.5m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(1, (float)(18.5m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(2, (float)(40.0m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(3, (float)(15.0m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(4, (float)(15.0m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(5, (float)(15.0m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(6, (float)(15.0m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(7, (float)(18.0m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(8, (float)(45.0m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(9, (float)(15.5m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(10, (float)(15.1m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(11, (float)(15.1m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(12, (float)(15.1m * RtfConstants.MILLIMETERS_TO_POINTS));
                table.setColWidth(13, (float)(15.1m * RtfConstants.MILLIMETERS_TO_POINTS));

                table.cell(0, 0).addParagraph().Text = "№ п/п";
                table.cell(0, 1).addParagraph().Text = "Рег.номер";
                table.cell(0, 2).addParagraph().Text = "ФИО";
                table.cell(0, 3).addParagraph().Text = "Сумма баллов";
                table.cell(0, 4).addParagraph().Text = "Проф. экзамен";
                table.cell(0, 5).addParagraph().Text = "Доп. экзамен";
                table.cell(0, 6).addParagraph().Text = "Конкурс";
                table.cell(0, 7).addParagraph().Text = "Подлинники";
                table.cell(0, 8).addParagraph().Text = "Контакты";
                table.cell(0, 9).addParagraph().Text = "Медалист";
                table.cell(0, 10).addParagraph().Text = "Серия док. об обр.";                
                table.cell(0, 11).addParagraph().Text = "ср. балл";
                table.cell(0, 12).addParagraph().Text = "Ретинг. коэфф.";
                if (dgvAbits.Columns.Contains("Олимпиада"))
                    table.cell(0, 13).addParagraph().Text = "Олимпиада";

                for (int j = 0; j < 14; j++)
                {
                    // Устанавливаем горизонтальное и вертикальное выравнивание текста "по центру" в каждой ячейке таблицы
                    table.cell(0, j).Alignment = Align.Center;
                    table.cell(0, j).AlignmentVertical = AlignVertical.Middle;
                }

                int r = 0;
                foreach (DataGridViewRow row in dgvAbits.Rows)
                {
                    ++r;
                    table.cell(r, 0).addParagraph().Text = r.ToString();
                    table.cell(r, 1).addParagraph().Text = row.Cells["Рег_Номер"].Value.ToString();
                    table.cell(r, 2).addParagraph().Text = row.Cells["ФИО"].Value.ToString();
                    table.cell(r, 3).addParagraph().Text = row.Cells["Сумма баллов"].Value.ToString();
                    table.cell(r, 4).addParagraph().Text = row.Cells["Проф. экзамен"].Value.ToString();
                    table.cell(r, 5).addParagraph().Text = row.Cells["Доп. экзамен"].Value.ToString();
                    table.cell(r, 6).addParagraph().Text = row.Cells["Конкурс"].Value.ToString();
                    table.cell(r, 7).addParagraph().Text = row.Cells["Подлинники документов"].Value.ToString();
                    table.cell(r, 8).addParagraph().Text = row.Cells["Контакты"].Value.ToString();
                    table.cell(r, 9).addParagraph().Text = row.Cells["Медалист"].Value.ToString();
                    table.cell(r, 10).addParagraph().Text = MainClass.dbType == PriemType.PriemMag ? row.Cells["Серия диплома"].Value.ToString() : row.Cells["Серия аттестата"].Value.ToString();
                    table.cell(r, 11).addParagraph().Text = row.Cells["Средний балл"].Value.ToString();
                    table.cell(r, 12).addParagraph().Text = row.Cells["Рейтинговый коэффициент"].Value.ToString();
                    if(dgvAbits.Columns.Contains("Олимпиада"))
                        table.cell(r, 13).addParagraph().Text = row.Cells["Олимпиада"].Value.ToString(); 

                    for (int j = 0; j < 14; j++)
                    {
                        // Устанавливаем горизонтальное и вертикальное выравнивание текста "по центру" в каждой ячейке таблицы
                        table.cell(r, j).Alignment = Align.Center;
                        table.cell(r, j).AlignmentVertical = AlignVertical.Middle;
                    }
                }

                // Задаём толщину внутренних границ таблицы
                table.setInnerBorder(RtfWriter.BorderStyle.Single, 0.5f);
                // Задаём толщину внешних границ таблицы
                table.setOuterBorder(RtfWriter.BorderStyle.Single, 0.5f);
                
                doc.save(MainClass.saveTempFolder + "\\RatingList.rtf");

                // ==========================================================================
                // Открываем сохранённый RTF файл
                // ==========================================================================
                WordDoc wd = new WordDoc(string.Format(@"{0}\RatingList.rtf", MainClass.saveTempFolder));
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при составлении списка:\n" + ex.Message +
                    ex.InnerException == null ? "" : ("\nВнутреннее исключение:\n" + ex.InnerException.Message));
            }
        }

        private void btnLock_Click(object sender, EventArgs e)
        {
            LockUnlock(true);
        }

        private void LockUnlock(bool locked)
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    context.FixierenView_UpdateLocked(MainClass.studyLevelGroupId, FacultyId, LicenseProgramId, ObrazProgramId, ProfileId, StudyBasisId, StudyFormId, IsSecond, IsReduced, IsParallel, IsCel, locked);
                    
                    lblLocked.Text = locked ? "ЗАЛОЧЕНА" : "НЕ залочена";
                }
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при локе/анлоке");
            }
            return;            
        }

        private void btnUnlock_Click(object sender, EventArgs e)
        {
            LockUnlock(false);
        }

        private void btnFixierenWeb_Click(object sender, EventArgs e)
        {
            WebFixieren();
        }

        private void WebFixieren()
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew))
                    {
                        Guid? fixViewId =
                            (from fv in context.FixierenView
                             where fv.StudyLevelGroupId == MainClass.studyLevelGroupId && fv.IsReduced == IsReduced && fv.IsParallel == IsParallel && fv.IsSecond == IsSecond
                             && fv.FacultyId == FacultyId && fv.LicenseProgramId == LicenseProgramId
                             && fv.ObrazProgramId == ObrazProgramId
                             && (ProfileId == null ? fv.ProfileId == null : fv.ProfileId == ProfileId)
                             && fv.StudyFormId == StudyFormId
                             && fv.StudyBasisId == StudyBasisId
                             && fv.IsCel == IsCel
                             && fv.IsCrimea == IsCrimea
                             && fv.IsQuota == IsQuota
                             select fv.Id).FirstOrDefault();

                        Guid? entryId =
                            (from fv in context.qEntry
                             where fv.StudyLevelGroupId == MainClass.studyLevelGroupId && fv.IsReduced == IsReduced && fv.IsParallel == IsParallel && fv.IsSecond == IsSecond
                             && fv.FacultyId == FacultyId && fv.LicenseProgramId == LicenseProgramId
                             && fv.ObrazProgramId == ObrazProgramId
                             && (ProfileId == null ? fv.ProfileId == null : fv.ProfileId == ProfileId)
                             && fv.StudyFormId == StudyFormId
                             && fv.StudyBasisId == StudyBasisId
                             select fv.Id).FirstOrDefault();
                        
                        //удалили старое
                        context.FirstWave_DELETE(entryId, IsCel, IsCrimea);

                        var fix = from fx in context.Fixieren
                                  where fx.FixierenViewId == fixViewId
                                  select fx;

                        //foreach(Fixieren f in fix)
                        int cnt = 0;                    
                        foreach (DataGridViewRow row in dgvAbits.Rows)                        
                        {
                            cnt++;
                            Guid? abId = new Guid(row.Cells["Id"].Value.ToString());
                            if (!chbCel.Checked)
                            {
                                if (!IsCrimea && !IsQuota)
                                    context.FirstWave_INSERT(abId, cnt);
                                else if (IsCrimea)
                                    context.FirstWave_INSERTCRIMEA(abId, cnt);
                                else if (IsQuota)
                                    context.FirstWave_INSERTQUOTA(abId, cnt);

                            }
                            else
                                context.FirstWave_INSERTCEL(abId, cnt);
                            //context.FirstWave_INSERT(f.AbiturientId, f.Number);
                        }
                        transaction.Complete();
                    }
                }
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при WEB FIXIEREN !");
            }
            MessageBox.Show("DONE!");
        }        

        private void btnUnfix_Click(object sender, EventArgs e)
        {
            Unfixieren();
        }

        private void Unfixieren()
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    Guid? entryId = (from fv in context.qEntry
                                         where fv.StudyLevelGroupId == MainClass.studyLevelGroupId && fv.IsReduced == IsReduced && fv.IsParallel == IsParallel && fv.IsSecond == IsSecond
                                         && fv.FacultyId == FacultyId && fv.LicenseProgramId == LicenseProgramId
                                         && fv.ObrazProgramId == ObrazProgramId
                                         && (ProfileId == null ? fv.ProfileId == null : fv.ProfileId == ProfileId)
                                         && fv.StudyFormId == StudyFormId
                                         && fv.StudyBasisId == StudyBasisId
                                         select fv.Id).FirstOrDefault();
                    
                    //удалили
                    context.FirstWave_DELETE(entryId, IsCel, IsCrimea);
                }
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при WEB FIXIEREN !");
            }
            MessageBox.Show("DONE!");
        }

        private void btnDeleteAb_Click(object sender, EventArgs e)
        {
            if (MainClass.IsPasha())
            {
                using (PriemEntities context = new PriemEntities())
                {
                    if (MessageBox.Show("Удалить из рейтингового списка?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew))
                        {
                            foreach (DataGridViewRow dgvr in dgvAbits.SelectedRows)
                            {
                                Guid abId = new Guid(dgvr.Cells["Id"].Value.ToString());
                                try
                                {
                                    context.Fixieren_DELETE(abId);
                                    context.FirstWave_DeleteByAbId(abId);
                                }
                                catch (Exception ex)
                                {
                                    WinFormsServ.Error("Ошибка удаления данных" + ex.Message);
                                }
                            }

                            transaction.Complete();
                        }   
                        UpdateDataGrid();
                    }
                }
            }
        }

        private void btnUpdateGrid_Click(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }

        private void chbIsReduced_CheckedChanged(object sender, EventArgs e)
        {
            FillStudyForm();
            NullDataGrid();
        }
        private void chbIsParallel_CheckedChanged(object sender, EventArgs e)
        {
            FillStudyForm();
            NullDataGrid();
        }
        private void chbIsSecond_CheckedChanged(object sender, EventArgs e)
        {
            FillStudyForm();
            NullDataGrid();
        }

        private void chbCel_CheckedChanged(object sender, EventArgs e)
        {
            NullDataGrid();

            if (IsQuota)
                chbIsQuota.Checked = false;
            if (IsCrimea)
                chbIsCrimea.Checked = false;
            //if (IsCel)
            //    btnFixierenWeb.Enabled = false;
            //else
            //    btnFixierenWeb.Enabled = true;
        }
        private void chbIsCrimea_CheckedChanged(object sender, EventArgs e)
        {
            NullDataGrid();
            if (IsCel)
                chbCel.Checked = false;
            if (IsQuota)
                chbIsQuota.Checked = false;
        }
        private void chbIsQuota_CheckedChanged(object sender, EventArgs e)
        {
            NullDataGrid();
            if (IsCel)
                chbCel.Checked = false;
            if (IsCrimea)
                chbIsCrimea.Checked = false;
        }

        private void btnToExcel_Click(object sender, EventArgs e)
        {
            PrintClass.PrintAllToExcel(dgvAbits);
        }             
    }
}