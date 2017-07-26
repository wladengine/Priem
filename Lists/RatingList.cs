using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Transactions;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using EducServLib;
using BDClassLib;
using WordOut;

using RtfWriter;
using PriemLib;
using System.Data.Entity.Core.Objects;
using System.Threading.Tasks;

namespace Priem
{
    public partial class RatingList : BookList
    {
        string _queryFrom;
        string _queryBody;
        string _queryOlymps;
        string _queryOrange;

        bool bMagAddNabor1Enabled = MainClass.bMagAddNabor1Enabled;
        DateTime dtMagAddNabor1 = MainClass.dtMagAddNabor1;

        bool b1kursAddNabor1Enabled = MainClass.b1kursAddNabor1Enabled;
        DateTime dt1kursAddNabor1 = MainClass.dt1kursAddNabor1;

        bool bFirstWaveEnabled = MainClass.bFirstWaveEnabled;

        BackgroundWorker bw = new BackgroundWorker();

        //constructor
        public RatingList(bool fromFixieren)
        {
            InitializeComponent();
            InitVariables();

            bw.WorkerSupportsCancellation = true;
            bw.DoWork += UpdateDataGrid_DoWorkAsync;
            bw.RunWorkerCompleted += UpdateDataGrid_AsyncWorkCompleted;

            _queryBody = @"SELECT DISTINCT qAbiturient.Id as Id, qAbiturient.RegNum as Рег_Номер, 
                    extPerson.PersonNum as 'Ид. номер', qAbiturient.Priority as [Приоритет], extPerson.FIO as ФИО, 
                    qAbiturient.Sum + extAbitAdditionalMarksSum.AdditionalMarksSum as 'Сумма баллов', 
                    qAbiturient.Sum as 'Сумма баллов (осн)', 
                    extAbitAdditionalMarksSum.AdditionalMarksSum AS 'Сумма баллов (ИндДост)', 
                    qAbiturient.MarksCount as 'Кол-во оценок', 
                    case when extPerson.HasOriginals = 1 then 'Да' else 'Нет' end as 'Подлинники документов', 
                    qAbiturient.Coefficient as 'Рейтинговый коэффициент', 
                    Competition.Name as Конкурс, hlpAbiturientProf.Prof AS 'Проф. экзамен', 
                    hlpAbiturientProfAdd.ProfAdd AS 'Доп. экзамен',
                    Abiturient_FetchValues.MarkOrderNumber1 AS [Экзамен 1], Abiturient_FetchValues.MarkOrderNumber2 AS [Экзамен 2], Abiturient_FetchValues.MarkOrderNumber3 AS [Экзамен 3],
                    CASE WHEN EXISTS(SELECT Id FROM ed.hlpProfOlympiads AS Olympiads WHERE OlympValueId = 6 AND AbiturientId = qAbiturient.Id) THEN 1 
                        else CASE WHEN EXISTS(SELECT Id FROM ed.hlpProfOlympiads AS Olympiads WHERE OlympValueId = 5 AND AbiturientId = qAbiturient.Id) THEN 2 
                            else CASE WHEN EXISTS(SELECT Id FROM ed.hlpProfOlympiads AS Olympiads WHERE OlympValueId = 7 AND AbiturientId = qAbiturient.Id) THEN 3 
                                else 4 
                                END 
                            END 
                        END AS olymp,
                    CASE WHEN extPerson_EducationInfo_Current.AttestatSeries IN ('ЗА','ЗБ','ЗВ','АЗ') then 1 else CASE WHEN extPerson_EducationInfo_Current.AttestatSeries IN ('СА','СБ','СВ') then 2 else 3 end end as attestat,
                    (CASE WHEN extPerson_EducationInfo_Current.IsExcellent=1 THEN 5 ELSE extPerson_EducationInfo_Current.SchoolAVG END) as attAvg,
                    CASE WHEN (CompetitionId=1  OR CompetitionId=8) then 1 else case when (CompetitionId=2 OR CompetitionId=7) AND extPerson.Privileges>0 then 2 else (case when CompetitionId=6 then 3 else 4 end) end end as comp, 
                    (CASE WHEN CompetitionId NOT IN (1, 8) 
                     THEN 0 ELSE (CASE WHEN hlpAbiturient_Olympiads_SortLevel1.[PersonId] IS NOT NULL 
                                              THEN 100 + extAbitAdditionalMarksSum.AdditionalMarksSum ELSE (CASE WHEN hlpAbiturient_Olympiads_SortLevel2.[PersonId] IS NOT NULL 
                                              THEN 90 + extAbitAdditionalMarksSum.AdditionalMarksSum ELSE (CASE WHEN hlpAbiturient_Olympiads_SortLevel3.[PersonId] IS NOT NULL 
                                              THEN 80 + extAbitAdditionalMarksSum.AdditionalMarksSum ELSE (CASE WHEN hlpAbiturient_Olympiads_SortLevel4.[PersonId] IS NOT NULL 
                                              THEN 70 + extAbitAdditionalMarksSum.AdditionalMarksSum ELSE 50 END) END) END) END) END) as noexamssort,
                    (CASE WHEN CompetitionId NOT IN (1, 8) THEN 0 ELSE extAbitAdditionalMarksSum.AdditionalMarksSum END) AS noexamsKoefsort,
                    (CASE WHEN CompetitionId NOT IN (1, 8) THEN 0 ELSE extPerson_EducationInfo_Current.SchoolAVG END) AS noexamsAttAVGSort,
                    (CASE WHEN CompetitionId NOT IN (1, 8) THEN 0 ELSE (CASE WHEN extPerson.Privileges > 0 THEN 10 ELSE 1 END) END) AS noexamsPrivSort,
                    CASE WHEN (CompetitionId=5 OR CompetitionId=9) then 1 else 0 end as preimsort,
                    case when extPerson_EducationInfo_Current.IsExcellent>0 then 'Да' else 'Нет' end as 'Медалист', 
                    extPerson_EducationInfo_Current.AttestatSeries as 'Серия аттестата', 
                    extPerson_EducationInfo_Current.DiplomSeries as 'Серия диплома', 
                    extPerson_EducationInfo_Current.SchoolAVG as 'Средний балл', 
                    extPerson.Email + ', '+ extPerson.Phone + ', ' + extPerson.Mobiles AS 'Контакты',
                    hlpAbiturientProf.Prof AS ProfSort, hlpAbiturientProfAdd.ProfAdd
                    /* (CASE WHEN hlpEntryWithAddExams.EntryId IS NULL THEN hlpAbiturientProf.Prof ELSE hlpAbiturientProfAdd.ProfAdd END) AS DopOrProfSort */";

            _queryOlymps = MainClass.lstStudyLevelGroupId.First() == 1 ? @", (SELECT TOP(1) extOlympiads.OlympValueAcr + '-' + extOlympiads.OlympName FROM ed.extOlympiadsAll AS extOlympiads 
                           WHERE extOlympiads.AbiturientId = qAbiturient.Id AND extOlympiads.OlympTypeId = 3 order by extOlympiads.sortOrder) as 'Олимпиада' " : "";

            _queryFrom = @" FROM ed.qAbiturientAll AS qAbiturient
INNER JOIN ed.extPerson ON extPerson.Id = qAbiturient.PersonId
INNER JOIN ed.extPerson_EducationInfo_Current ON extPerson_EducationInfo_Current.PersonId = extPerson.Id
INNER JOIN ed.Competition ON Competition.Id = qAbiturient.CompetitionId 
INNER JOIN ed.extEnableProtocol ON extEnableProtocol.AbiturientId = qAbiturient.Id 
LEFT JOIN ed.hlpEntryWithAddExams ON hlpEntryWithAddExams.EntryId = qAbiturient.EntryId
LEFT JOIN ed.hlpAbiturientProfAdd ON hlpAbiturientProfAdd.Id = qAbiturient.Id 
LEFT JOIN ed.hlpAbiturientProf ON hlpAbiturientProf.Id = qAbiturient.Id 
LEFT JOIN ed.extAbitAdditionalMarksSum ON qAbiturient.Id = extAbitAdditionalMarksSum.AbiturientId
LEFT JOIN ed.Abiturient_FetchValues ON qAbiturient.Id = Abiturient_FetchValues.AbiturientId
LEFT JOIN ed.hlpMinMarkAbiturient ON hlpMinMarkAbiturient.Id = qAbiturient.Id
LEFT JOIN ed.hlpAbiturient_Olympiads_SortLevel1 ON qAbiturient.PersonId = hlpAbiturient_Olympiads_SortLevel1.[PersonId] 
LEFT JOIN ed.hlpAbiturient_Olympiads_SortLevel2 ON qAbiturient.PersonId = hlpAbiturient_Olympiads_SortLevel2.[PersonId] 
LEFT JOIN ed.hlpAbiturient_Olympiads_SortLevel3 ON qAbiturient.PersonId = hlpAbiturient_Olympiads_SortLevel3.[PersonId] 
LEFT JOIN ed.hlpAbiturient_Olympiads_SortLevel4 ON qAbiturient.PersonId = hlpAbiturient_Olympiads_SortLevel4.[PersonId]
LEFT JOIN ed._FirstWaveBackUp FW ON FW.AbiturientId = qAbiturient.Id";

            if (MainClass.dbType == PriemType.PriemMag)
                _queryFrom += " LEFT JOIN ed.hlpMinMarkMag ON hlpMinMarkMag.AbiturientId = qAbiturient.Id";

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
            FillStudyLevelGroup();
            FillStudyForm();
            FillLicenseProgram();
            FillObrazProgram();
            FillProfile();

            //если 
            //chbCel.Visible = false;

            if (MainClass.dbType == PriemType.PriemMag)
                chbWithOlymps.Visible = false;
        }

        private void FillStudyLevelGroup()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = MainClass.GetEntry(context).Select(x => new { x.StudyLevelGroupId, x.StudyLevelGroupName }).ToList();

                List<KeyValuePair<string, string>> lst = ent.Select(u => new KeyValuePair<string, string>(u.StudyLevelGroupId.ToString(), u.StudyLevelGroupName)).Distinct().ToList();

                ComboServ.FillCombo(cbStudyLevelGroup, lst, false, false);
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

                var ent = MainClass.GetEntry(context).Where(c => c.FacultyId == FacultyId);

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
            cbStudyLevelGroup.SelectedIndexChanged += cbStudyLevelGroup_SelectedIndexChanged;
        }

        void cbStudyLevelGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillStudyForm();
            NullDataGrid();
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
        }

        private void btnUpdateGrid_Click(object sender, EventArgs e)
        {
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
                int plan = 0, planCel = 0, planCrimea = 0, planQuota = 0, entered = 0, enteredCel = 0, enteredQuota = 0;               

                qEntry entry = (from ent in MainClass.GetEntry(context)
                       where ent.IsReduced == IsReduced && ent.IsParallel == IsParallel && ent.IsSecond == IsSecond 
                       && ent.FacultyId == FacultyId && ent.LicenseProgramId == LicenseProgramId
                       && ent.ObrazProgramId == ObrazProgramId
                       && (ProfileId == null ? ent.ProfileId == 0 : ent.ProfileId == ProfileId)
                       && ent.StudyFormId == StudyFormId
                       && ent.StudyBasisId == StudyBasisId
                       select ent).FirstOrDefault();

                if (entry == null)
                    return 0;

                plan = entry.KCP ?? 0;
                planCel = entry.KCPCel ?? 0;
                planQuota = entry.KCPQuota ?? 0;

                Guid? entryId = entry.Id;

                entered = (from ab in context.qAbitAll
                           join ev in context.extEntryView
                           on ab.Id equals ev.AbiturientId
                           where ab.CompetitionId != 6 && ab.CompetitionId != 11 && ab.CompetitionId != 12 && ab.CompetitionId != 2 && ab.CompetitionId != 7 && ab.EntryId == entryId
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

                planCrimea = context.Entry.Where(x => x.ParentEntryId == entryId).Select(x => x.KCP).DefaultIfEmpty(0).First() ?? 0;
               
                CheckLockAndPasha(context);

                if (IsCel)
                    return planCel - enteredCel;
                else if (IsQuota)
                    return planQuota - enteredQuota;
                else
                {
                    return plan - enteredCel - entered - enteredQuota;
                }
            }
        }

        private void CheckLockAndPasha(PriemEntities context)
        {
            //лочит кнопку 
            FixierenView fixView =
                (from fv in context.FixierenView
                 where fv.StudyLevelGroupId == StudyLevelGroupId
                 && fv.IsReduced == IsReduced && fv.IsParallel == IsParallel && fv.IsSecond == IsSecond
                 && fv.FacultyId == FacultyId && fv.LicenseProgramId == LicenseProgramId
                 && fv.ObrazProgramId == ObrazProgramId
                 && (ProfileId == null ? fv.ProfileId == 0 : fv.ProfileId == ProfileId)
                 && fv.StudyFormId == StudyFormId
                 && fv.StudyBasisId == StudyBasisId
                 && fv.IsCel == IsCel
                 && fv.IsQuota == IsQuota
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
            if (bw.IsBusy)
            {
                bw.CancelAsync();
                return;
            }

            if (!StudyFormId.HasValue || !StudyBasisId.HasValue || !FacultyId.HasValue || !LicenseProgramId.HasValue || !ObrazProgramId.HasValue)
                return;

            try
            {                
                string sOrderBy = string.Empty;
                if (MainClass.dbType == PriemType.PriemMag)
                {
                    sOrderBy =
                        chbCel.Checked ?
                        " ORDER BY qAbiturient.Coefficient, comp , noexamssort desc, 'Сумма баллов' desc, qAbiturient.MarksCount desc, ФИО" :
                        " ORDER BY comp , 'Сумма баллов' desc, 'Проф. экзамен' DESC, qAbiturient.Coefficient DESC, /*attAvg desc,*/ [Средний балл] desc, qAbiturient.MarksCount desc, ФИО";
                }
                else
                {
                    sOrderBy =
                        chbCel.Checked ?
                        " ORDER BY qAbiturient.Coefficient, comp, noexamssort desc, 'Сумма баллов' desc, 'Сумма баллов (осн)' DESC, ProfSort desc, ProfAdd desc, qAbiturient.MarksCount desc, ФИО"
                        :
                        " ORDER BY comp, noexamssort DESC, noexamsKoefsort DESC, noexamsPrivSort, noexamsAttAVGSort DESC, 'Сумма баллов' desc, 'Сумма баллов (осн)' DESC, [Экзамен 1] desc, [Экзамен 2] desc, [Экзамен 3] desc, preimsort desc, ProfAdd desc, " +
                        "olymp, Медалист, attAvg desc, qAbiturient.Coefficient, qAbiturient.MarksCount desc, ФИО"
                        ;
                }
                string totalQuery = null;

                
                plan = GetPlanValueAndCheckLock();

                if (chbFix.Checked)
                {
                    if (MainClass.dbType == PriemType.PriemMag)
                        _queryOrange = @", CASE WHEN EXISTS(SELECT PersonId FROM ed.hlpPersonsWithOriginals WHERE PersonId = qAbiturient.PersonId AND EntryId <> qAbiturient.EntryId) then 1 else 0 end as orange ";
                    else
                        _queryOrange = @", CASE WHEN EXISTS(SELECT extEntryView.Id FROM ed.extEntryView INNER JOIN ed.Abiturient a ON extEntryView.AbiturientId = a.Id WHERE a.PersonId = qAbiturient.PersonId) then 1 else 0 end as orange ";

                    string queryFix = _queryBody + _queryOrange +
                    @" FROM ed.qAbiturient 
                    INNER JOIN ed.extPerson ON extPerson.Id = qAbiturient.PersonId                    
                    INNER JOIN ed.extPerson_EducationInfo_Current ON extPerson_EducationInfo_Current.PersonId = extPerson.Id
                    INNER JOIN ed.Competition ON Competition.Id = qAbiturient.CompetitionId 
                    INNER JOIN ed.Fixieren ON Fixieren.AbiturientId = qAbiturient.Id 
                    LEFT JOIN ed.hlpEntryWithAddExams ON hlpEntryWithAddExams.EntryId = qAbiturient.EntryId
                    LEFT JOIN ed.FixierenView ON Fixieren.FixierenViewId = FixierenView.Id 
                    LEFT JOIN ed.hlpAbiturientProfAdd ON hlpAbiturientProfAdd.Id = qAbiturient.Id 
                    LEFT JOIN ed.hlpAbiturientProf ON hlpAbiturientProf.Id = qAbiturient.Id 

                    LEFT JOIN ed.extAbitAdditionalMarksSum ON qAbiturient.Id = extAbitAdditionalMarksSum.AbiturientId
                    LEFT JOIN ed.Abiturient_FetchValues ON qAbiturient.Id = Abiturient_FetchValues.AbiturientId
                    LEFT JOIN ed.hlpAbiturient_Olympiads_SortLevel1 ON qAbiturient.Id = hlpAbiturient_Olympiads_SortLevel1.[AbiturientId] 
                    LEFT JOIN ed.hlpAbiturient_Olympiads_SortLevel2 ON qAbiturient.Id = hlpAbiturient_Olympiads_SortLevel2.[AbiturientId] 
                    LEFT JOIN ed.hlpAbiturient_Olympiads_SortLevel3 ON qAbiturient.Id = hlpAbiturient_Olympiads_SortLevel3.[AbiturientId] 
                    LEFT JOIN ed.hlpAbiturient_Olympiads_SortLevel4 ON qAbiturient.Id = hlpAbiturient_Olympiads_SortLevel4.[AbiturientId]
";

                    string whereFix = string.Format(@"
WHERE FixierenView.StudyLevelGroupId IN ({10}) AND FixierenView.StudyFormId={0} AND FixierenView.StudyBasisId={1} AND FixierenView.FacultyId={2} 
AND FixierenView.LicenseProgramId={3} AND FixierenView.ObrazProgramId={4} AND FixierenView.ProfileId='{5}' AND FixierenView.IsCel = {6}
AND FixierenView.IsSecond = {7} AND FixierenView.IsReduced = {8} AND FixierenView.IsParallel = {9} AND FixierenView.IsQuota = {11}",
                        StudyFormId, StudyBasisId, FacultyId, LicenseProgramId, ObrazProgramId, ProfileId,
                        QueryServ.StringParseFromBool(IsCel), QueryServ.StringParseFromBool(IsSecond), QueryServ.StringParseFromBool(IsReduced), QueryServ.StringParseFromBool(IsParallel), 
                        Util.BuildStringWithCollection(MainClass.lstStudyLevelGroupId), QueryServ.StringParseFromBool(IsQuota));

                    //whereFix += " AND Fixieren.AbiturientId NOT IN (SELECT AbiturientId FROM ed.extEntryView) AND qAbiturient.BackDoc = 0 ";
                    //sOrderBy = " ORDER BY Fixieren.Number ";

                    totalQuery = queryFix + whereFix + sOrderBy;
                }
                else
                {
                    string sFilters = GetFilterString();
                    
                    //целевики?
                    if (chbCel.Checked)
                        sFilters += " AND qAbiturient.CompetitionId = 6";
                    // в общем списке выводить всех 
                    else
                        sFilters += " AND qAbiturient.CompetitionId <> 6";

                    //не забрали доки
                    sFilters += " AND qAbiturient.BackDoc = 0 ";

                    sFilters += " AND qAbiturient.Id NOT IN (SELECT AbiturientId FROM ed.extEntryView) ";

                    //не иностранцы
                    sFilters += " AND qAbiturient.IsForeign = 0 ";

                    //квотники?
                    if (IsQuota)
                        sFilters += " AND qAbiturient.CompetitionId IN (2, 7) ";
                    else
                        sFilters += " AND qAbiturient.CompetitionId NOT IN (2, 7) ";

                    // кроме бэ преодолены мин планки 
                    if (MainClass.dbType == PriemType.PriemMag)
                        sFilters += " AND ((CompetitionId=1 OR CompetitionId=8) OR hlpMinMarkMag.AbiturientId IS NULL)";
                    else
                        sFilters += " AND ((CompetitionId=1 OR CompetitionId=8) OR hlpMinMarkAbiturient.Id IS NULL)";

                    string examsCnt = _bdc.GetStringValue(string.Format(" SELECT Count(Id) FROM ed.extExamInEntry WHERE EntryId='{0}' AND ParentExamInEntryBlockId IS NULL AND extExamInEntry.ExamId <> 850", EntryId.ToString()));
                   
                    if (MainClass.dbType == PriemType.PriemMag)
                    { 
                        _queryOrange = @", CASE WHEN EXISTS(SELECT PersonId FROM ed.hlpPersonsWithOriginals WHERE PersonId = qAbiturient.PersonId AND EntryId <> qAbiturient.EntryId) then 1 else 0 end as orange ";

                        // кроме бэ нужное кол-во оценок есть
                        sFilters += " AND ((CompetitionId=1 OR CompetitionId=8) OR qAbiturient.MarksCount = " + examsCnt + " ) ";

                        if (bMagAddNabor1Enabled)
                            sFilters += " AND qAbiturient.DocInsertDate > '" + dtMagAddNabor1.ToShortDateString() + "' ";

                        totalQuery = _queryBody + _queryOrange + _queryFrom + sFilters + sOrderBy;
                    }
                    else
                    {
                        _queryOrange = @", CASE WHEN EXISTS(SELECT extEntryView.Id FROM ed.extEntryView INNER JOIN ed.Abiturient a ON extEntryView.AbiturientId = a.Id WHERE a.PersonId = qAbiturient.PersonId) then 1 else 0 end as orange ";

                        if (bFirstWaveEnabled && MainClass.dbType == PriemType.Priem && StudyBasisId != 2)
                            sFilters += " AND FW.AbiturientId IS NOT NULL";

                        // кроме бэ нужное кол-во оценок есть
                        sFilters += " AND ((CompetitionId=1 OR CompetitionId=8) OR qAbiturient.MarksCount = " + examsCnt + " ) ";

                        if (MainClass.dbType == PriemType.Priem && b1kursAddNabor1Enabled)
                            sFilters += " AND qAbiturient.DocInsertDate > '" + dt1kursAddNabor1.ToShortDateString() + "' ";

                        //до зачисления льготников выводить их, а потом - убирать
                        //sFilters += " AND ed.qAbiturient.CompetitionId NOT IN (6, 1, 2, 7) ";                                        

                        // кроме бэ и тех, у кого нет сертификатов и оценок нужное кол-во оценок есть
                        sFilters += @"
                            AND 
                            (
                                CompetitionId IN (1, 8) 
                                OR 
                                (
                                    qAbiturient.PersonId NOT IN (SELECT PersonId FROM ed.EgeCertificate) 
                                    AND qAbiturient.Id NOT IN (SELECT AbiturientId FROM ed.Mark WHERE IsFromEge = 1) 
                                    AND qAbiturient.IsSecond = 0 AND qAbiturient.IsReduced = 0 AND qAbiturient.IsParallel = 0
                                ) 
                                OR qAbiturient.MarksCount = 
                                (
		                            SELECT COUNT(*) 
		                            FROM ed.extExamInEntry 
		                            WHERE extExamInEntry.EntryId = qAbiturient.EntryId --AND extExamInEntry.ExamId = 850
	                            )
                            ) ";
                        
                        //if (StudyBasisId == 2)
                        //    sFilters += " AND qAbiturient.Id NOT IN (SELECT AbiturientId FROM ed._FirstWaveGreen)";

                        totalQuery = _queryBody + (chbWithOlymps.Checked ? _queryOlymps : "") + _queryOrange + _queryFrom + sFilters + sOrderBy;
                    }
                }

                if (!dgvAbits.Columns.Contains("Number"))
                    dgvAbits.Columns.Add("Number", "№ п/п");

                lblCount.Text = "             Cвободных мест: " + plan;

                bw.RunWorkerAsync(new { dgv = dgvAbits, _bdc = _bdc, _sQuery = totalQuery, filters = "", _orderBy = "" });
                gbWait.Visible = true;
                UpdateControlsEnableStatus(false);
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при обновлении списка.", ex);
            }
        }

        async void UpdateDataGrid_DoWorkAsync(object sender, DoWorkEventArgs e)
        {
            SQLClass BDC = new SQLClass();
            BDC.OpenDatabase(MainClass.connString);

            Task<DataView> task = HelpClass.GetDataViewAsync((DataGridView)((dynamic)e.Argument).dgv, /*(BDClass)((dynamic)e.Argument)._bdc*/BDC, (string)((dynamic)e.Argument)._sQuery, (string)((dynamic)e.Argument).filters, _orderBy, false);

            while (!task.IsCompleted)
            {
                if (task.IsFaulted)
                {
                    e.Cancel = true;
                    return;
                }
                if (bw.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }

                System.Threading.Thread.Sleep(25);
            }

            if (task.IsFaulted)
                e.Cancel = true;
            else
                e.Result = await task;
        }
        void UpdateDataGrid_AsyncWorkCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            gbWait.Visible = false;
            UpdateControlsEnableStatus(true);

            if (e.Cancelled || (e.Error != null))
                return;

            dgvAbits.DataSource = e.Result;

            dgvAbits.Columns["Id"].Visible = false;
            dgvAbits.Columns["comp"].Visible = false;
            dgvAbits.Columns["noexamssort"].Visible = false;
            dgvAbits.Columns["noexamsKoefsort"].Visible = false;
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

            lblCount.Text = "Всего: " + dgvAbits.RowCount.ToString() + "             Cвободных мест: " + plan;
        }

        private void UpdateControlsEnableStatus(bool status)
        {
            cbFaculty.Enabled = status;
            cbLicenseProgram.Enabled = status;
            cbObrazProgram.Enabled = status;
            cbProfile.Enabled = status;
            cbStudyBasis.Enabled = status;
            cbStudyForm.Enabled = status;
            cbStudyLevelGroup.Enabled = status;
            tbFIO.Enabled = status;
            tbNumber.Enabled = status;
            btnCard.Enabled = status;
            btnAdd.Enabled = status;
            btnDeleteAb.Enabled = status;
            btnFixieren.Enabled = status;
            btnFixierenWeb.Enabled = status;
            btnLock.Enabled = status;
            btnRemove.Enabled = status;
            btnToExcel.Enabled = status;
            btnUnfix.Enabled = status;
            btnUnlock.Enabled = status;
            btnWord.Enabled = status;

            chbCel.Enabled = status;
            chbFix.Enabled = status;
            chbIsParallel.Enabled = status;
            chbIsQuota.Enabled = status;
            chbIsReduced.Enabled = status;
            chbIsSecond.Enabled = status;
            chbWithOlymps.Enabled = status;

            btnUpdateGrid.Text = status ? "Вывести список" : "Отмена";
        }

        private string GetFilterString()
        {
            string s = " WHERE 1=1 ";
            s += " AND qAbiturient.StudyLevelGroupId = " + StudyLevelGroupId;  
            
            //s += " AND ed.qAbiturient.DocDate>='20120813'"; 

            //обработали факультет
            if (FacultyId != null)
                s += " AND qAbiturient.FacultyId = " + FacultyId;      
            
            //обработали форму обучения  
            if (StudyFormId != null)
                s += " AND qAbiturient.StudyFormId = " + StudyFormId;

            //обработали основу обучения  
            if (StudyBasisId != null)
                s += " AND qAbiturient.StudyBasisId = " + StudyBasisId;               

            //обработали Направление
            if (LicenseProgramId != null)
                s += " AND qAbiturient.LicenseProgramId = " + LicenseProgramId;

            //обработали Образ программу
            if (ObrazProgramId != null)
                s += " AND qAbiturient.ObrazProgramId = " + ObrazProgramId;

            //обработали профиль
            if (ProfileId != null)
                s += string.Format(" AND qAbiturient.ProfileId = '{0}'", ProfileId);
            else
                s += " AND qAbiturient.ProfileId = 0";


            s += " AND qAbiturient.IsSecond = " + (IsSecond ? " 1 " : " 0 ");
            s += " AND qAbiturient.IsReduced = " + (IsReduced ? " 1 " : " 0 ");
            s += " AND qAbiturient.IsParallel = " + (IsParallel ? " 1 " : " 0 ");

            if (chbCel.Checked)
                s += " AND qAbiturient.CompetitionId = 6 ";

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
                                           where fv.StudyLevelGroupId == StudyLevelGroupId && fv.IsReduced == IsReduced && fv.IsParallel == IsParallel && fv.IsSecond == IsSecond
                                           && fv.FacultyId == FacultyId && fv.LicenseProgramId == LicenseProgramId
                                           && fv.ObrazProgramId == ObrazProgramId
                                           && (ProfileId == null ? fv.ProfileId == 0 : fv.ProfileId == ProfileId)
                                           && fv.StudyFormId == StudyFormId
                                           && fv.StudyBasisId == StudyBasisId
                                           && fv.IsCel == IsCel
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
                        context.FixierenView_Insert(StudyLevelGroupId, FacultyId, LicenseProgramId, ObrazProgramId, ProfileId, StudyBasisId, StudyFormId, IsSecond, IsReduced, IsParallel, IsCel, rand, false, isQuota: IsQuota, id: fvId);
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
                        WinFormsServ.Error("Ошибка при сохранении списка", ex);
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
                Print.PrintRatingProtocol(StudyFormId, StudyBasisId, FacultyId, LicenseProgramId, ObrazProgramId, ProfileId, IsCel,
                    plan, sfd.FileName, IsSecond, IsReduced, IsParallel, IsQuota);
        }        

        private void btnWord_Click(object sender, EventArgs e)
        {
            new RaitingListToWord(dgvAbits.Columns.Contains("Олимпиада"), dgvAbits).Show();
        }

        private void btnLock_Click(object sender, EventArgs e)
        {
            LockUnlock(true);
        }
        private void btnUnlock_Click(object sender, EventArgs e)
        {
            LockUnlock(false);
        }
        private void LockUnlock(bool locked)
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    context.FixierenView_UpdateLocked(StudyLevelGroupId, FacultyId, LicenseProgramId, ObrazProgramId, ProfileId, StudyBasisId, StudyFormId, IsSecond, IsReduced, IsParallel, IsCel, isCrimea: false, locked: locked);
                    
                    lblLocked.Text = locked ? "ЗАЛОЧЕНА" : "НЕ залочена";
                }
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при локе/анлоке", ex);
            }
            return;
        }

        private void btnFixierenWeb_Click(object sender, EventArgs e)
        {
            WebFixieren();
        }
        private void WebFixieren()
        {
            int cnt = 0;
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew))
                    {
                        bool bIsForeign = MainClass.dbType == PriemType.PriemForeigners;
                        Guid? fixViewId =
                            (from fv in context.FixierenView
                             where fv.StudyLevelGroupId == StudyLevelGroupId && fv.IsReduced == IsReduced && fv.IsParallel == IsParallel && fv.IsSecond == IsSecond
                             && fv.FacultyId == FacultyId && fv.LicenseProgramId == LicenseProgramId
                             && fv.ObrazProgramId == ObrazProgramId
                             && (ProfileId == null ? fv.ProfileId == 0 : fv.ProfileId == ProfileId)
                             && fv.StudyFormId == StudyFormId
                             && fv.StudyBasisId == StudyBasisId
                             && fv.IsCel == IsCel
                             && fv.IsQuota == IsQuota
                             //&& fv.IsForeign == bIsForeign
                             select fv.Id).FirstOrDefault();

                        Guid? entryId =
                            (from fv in context.qEntry
                             where fv.StudyLevelGroupId == StudyLevelGroupId && fv.IsReduced == IsReduced && fv.IsParallel == IsParallel && fv.IsSecond == IsSecond
                             && fv.FacultyId == FacultyId && fv.LicenseProgramId == LicenseProgramId
                             && fv.ObrazProgramId == ObrazProgramId
                             && (ProfileId == null ? fv.ProfileId == 0 : fv.ProfileId == ProfileId)
                             && fv.StudyFormId == StudyFormId
                             && fv.StudyBasisId == StudyBasisId
                             && fv.IsForeign == bIsForeign
                             select fv.Id).FirstOrDefault();
                        
                        //удалили старое
                        context.FirstWave_DELETE(entryId, IsCel, IsQuota);

                        var fix = from fx in context.Fixieren
                                  where fx.FixierenViewId == fixViewId
                                  select fx;

                        //foreach(Fixieren f in fix)
                        
                        foreach (DataGridViewRow row in dgvAbits.Rows)                        
                        {
                            cnt++;
                            Guid? abId = new Guid(row.Cells["Id"].Value.ToString());
                            if (!chbCel.Checked)
                            {
                                if (!IsQuota)
                                    context.FirstWave_INSERT(abId, cnt);
                                else if (IsQuota)
                                    context.FirstWave_INSERTQUOTA(abId, cnt);
                            }
                            else
                                context.FirstWave_INSERTCEL(abId, cnt);
                            
                        }
                        transaction.Complete();
                    }
                }
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при WEB FIXIEREN !", ex);
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
                                     where fv.StudyLevelGroupId == StudyLevelGroupId && fv.IsReduced == IsReduced && fv.IsParallel == IsParallel && fv.IsSecond == IsSecond
                                     && fv.FacultyId == FacultyId && fv.LicenseProgramId == LicenseProgramId
                                     && fv.ObrazProgramId == ObrazProgramId
                                     && (ProfileId == null ? fv.ProfileId == 0 : fv.ProfileId == ProfileId)
                                     && fv.StudyFormId == StudyFormId
                                     && fv.StudyBasisId == StudyBasisId
                                     select fv.Id).FirstOrDefault();
                    
                    //удалили
                    context.FirstWave_DELETE(entryId, IsCel, IsQuota);
                }
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при WEB FIXIEREN!", ex);
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

        private void btnToExcel_Click(object sender, EventArgs e)
        {
            PrintClass.PrintAllToExcel(dgvAbits);
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            if (bw.IsBusy)
                bw.CancelAsync();

            base.OnClosing(e);
        }
    }
}