using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using System.Transactions;

using EducServLib;
using BDClassLib;
using PriemLib;

namespace Priem
{
    public partial class LoadEgeMarks : Form
    {
        private DBPriem bdc;
        private NewWatch wtc;
        private int marksCount;

        public LoadEgeMarks()
        {
            InitializeComponent();           
            InitControls();            
        }  

        //дополнительная инициализация контролов
        private void InitControls()
        {
            this.CenterToParent(); 
            this.MdiParent = MainClass.mainform;
            bdc = MainClass.Bdc;

            ComboServ.FillCombo(cbFaculty, HelpClass.GetComboListByTable("ed.qFaculty WHERE Id IN (SELECT FacultyId FROM ed.extEntry WHERE StudyLevelGroupId = 1)"), false, true);             
            FillExams();

            cbFaculty.SelectedIndexChanged += new EventHandler(cbFaculty_SelectedIndexChanged);

            UpdateGridAbits();
        }

        private void UpdateGridAbits()
        {
            string quer = @"SELECT DISTINCT extEntry.FacultyName AS [Факультет], COUNT(DISTINCT t.AbiturientId) AS [Абитуриентов], 
(
	SELECT CONVERT(nvarchar, MAX(extEnableProtocol.Date), 104) 
	+ ' ' + CONVERT(nvarchar, MAX(extEnableProtocol.Date), 108)
	FROM ed.extEnableProtocol WHERE extEnableProtocol.StudyLevelGroupId = 1 
	AND extEnableProtocol.FacultyId = extEntry.FacultyId
)  AS [Дата последнего протокола о допуске]
FROM
(
    SELECT Abiturient.Id AS AbiturientId, ExamInEntry.Id AS ExamInEntryId, Abiturient.EntryId
    FROM ed.Abiturient
    INNER JOIN ed.ExamInEntry ON ExamInEntry.EntryId = Abiturient.EntryId
    INNER JOIN ed.EgeToExam ON EgeToExam.ExamId = ExamInEntry.ExamId
    INNER JOIN ed.extEnableProtocol ON extEnableProtocol.AbiturientId = Abiturient.Id
    WHERE extEnableProtocol.IsOld = 0 AND extEnableProtocol.Excluded = 0
    AND Abiturient.BackDoc = 0 AND Abiturient.NotEnabled = 0
    AND Abiturient.CompetitionId NOT IN (1, 8)
    EXCEPT
    SELECT AbiturientId, ExamInEntryId, ExamInEntry.EntryId
    FROM ed.Mark
    INNER JOIN ed.ExamInEntry ON ExamInEntry.Id = Mark.ExamInEntryId
    INNER JOIN ed.EgeToExam ON EgeToExam.ExamId = ExamInEntry.ExamId
) t
INNER JOIN ed.extEntry ON extEntry.Id = t.EntryId
WHERE extEntry.StudyLevelGroupId = 1
GROUP BY extEntry.FacultyName, extEntry.FacultyId
ORDER BY 1";
            dgvProtocols.DataSource = MainClass.Bdc.GetDataSet(quer, new SortedList<string, object>() { { "@Date", MainClass._1k_LastEgeMarkLoadTime.AddMinutes(-10) } }).Tables[0];
            dgvProtocols.Columns["Факультет"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
            dgvProtocols.Columns["Абитуриентов"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
            dgvProtocols.Columns["Дата последнего протокола о допуске"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
        }

        void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillExams();
        }

        private void FillExams()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = Exams.GetExamsWithFilters(context, MainClass.lstStudyLevelGroupId, FacultyId, null, null, null, null, null, null, null, null).Where(c => !c.IsAdditional);               

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.ExamId.ToString(), u.ExamName)).Distinct().ToList();
                ComboServ.FillCombo(cbExam, lst, false, true);
            }
        }

        public int? FacultyId
        {
            get { return ComboServ.GetComboIdInt(cbFaculty); }
            set { ComboServ.SetComboId(cbFaculty, value); }
        }

        public int? ExamId
        {
            get { return ComboServ.GetComboIdInt(cbExam); }
            set { ComboServ.SetComboId(cbExam, value); }
        }

        //строим запрос фильтров для абитуриентов
        private string GetAbitFilterString(int iFacultyId)
        {
            string s = " AND E.StudyLevelGroupId = 1";
 
            //обработали факультет            
            if (iFacultyId != null)
                s += " AND E.FacultyId = " + iFacultyId.ToString(); 

            return s;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (!MainClass.IsPasha())
                return;
            
            LoadMarks();
        }

        private void LoadMarks()
        {
            if (!MainClass.IsPasha())
                return;

            try
            {
                wtc = new NewWatch(2);
                wtc.Show();
                marksCount = 0;

                if (FacultyId == null)
                {
                    using (PriemEntities context = new PriemEntities())
                    {
                        foreach (int facId in context.extEntry.Where(x => x.StudyLevelGroupId == 1).Select(x => x.FacultyId).Distinct().ToList())
                        {
                            var ent = Exams.GetExamsWithFilters(context, MainClass.lstStudyLevelGroupId, facId, null, null, null, null, null, null, null, null)
                                .Where(c => !c.IsAdditional && !c.IsSecond && !c.IsGosLine);
                            foreach (var exInEnt in ent.Select(x => x.ExamId).Distinct())
                                SetMarksForExam(exInEnt, facId);
                        }
                    }
                }
                else
                {
                    if (ExamId == null)
                    {
                        foreach (KeyValuePair<string, string> ex in cbExam.Items)
                        {
                            int exId;
                            if (int.TryParse(ex.Key, out exId))
                                SetMarksForExam(exId, FacultyId.Value);
                        }
                    }
                    else
                        SetMarksForExam(ExamId.Value, FacultyId.Value);
                }

                UpdateGridAbits();

                MainClass._1k_LastEgeMarkLoadTime = DateTime.Now;
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка загрузки оценок " + ex.Message);
            }
            finally
            {
                wtc.Close();
                MessageBox.Show(string.Format("Зачтено {0} оценок", marksCount));
            }

            
        }

        private void SetMarksForExam(int examId, int iFacultyId)
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    if (context.EgeToExam.Where(x => x.ExamId == examId).Count() == 0)
                        return;

                    int examInostr = 380;
                    int filFacId = 18;

                    string flt_backDoc = " AND qAbiturient.BackDoc = 0  ";
                    string flt_enable = " AND qAbiturient.NotEnabled = 0 ";
                    string flt_protocol = " AND ProtocolTypeId = 1 AND IsOld = 0 AND Excluded = 0 ";
                    string flt_status = " /*AND ed.extFBSStatus.FBSStatusId IN (1,4) */";
                    string flt_mark = string.Format(@" AND 
(
    qAbiturient.Id NOT IN 
    (
        /*оценка по НЕ-ДОПу ещё не проставлена*/
            SELECT Mark.AbiturientId 
            FROM ed.Mark 
            INNER JOIN ed.extExamInEntry ON Mark.ExamInEntryId = extExamInEntry.Id 
            WHERE Mark.AbiturientId = qAbiturient.Id 
            AND extExamInEntry.ExamId = {0}
            AND extExamInEntry.IsAdditional=0
    )
    /*OR qAbiturient.Id IN
    /*или у абитуриента зачёлся балл ниже, чем есть среди его ЕГЭ*/
    (
        SELECT qMark.AbiturientId
        FROM ed.Mark AS qMark
        INNER JOIN ed.Abiturient ABIT ON ABIT.Id = qMark.AbiturientId
        INNER JOIN ed.hlpEgeMarkMaxApprovedValue ON hlpEgeMarkMaxApprovedValue.PersonId = ABIT.PersonId 
        INNER JOIN ed.ExamInEntry ON ExamInEntry.Id = qMark.ExamInEntryId
        INNER JOIN ed.EgeToExam ON EgeToExam.ExamId = ExamInEntry.ExamId 
        WHERE EgeToExam.EgeExamNameId = hlpEgeMarkMaxApprovedValue.EgeExamNameId 
        AND qMark.IsFromEge = 1 
        AND qAbiturient.EntryId = ABIT.EntryId
        AND ExamInEntry.ExamId = EgeToExam.ExamId
        AND ExamInEntry.ExamId = {0}
        AND hlpEgeMarkMaxApprovedValue.EgeMarkValue > qMark.Value
    )*/
)", examId);
                    //string flt_hasEge = string.Format(" AND Person.Id IN (SELECT PersonId FROM ed.EgeMark LEFT JOIN ed.EgeToExam ON EgeMark.EgeExamNameId = EgeToExam.EgeExamNameId WHERE EgeToExam.ExamId = @ExamId)", examId);
                    string flt_hasExam = string.Format(" AND qAbiturient.EntryId IN (SELECT ed.ExamInEntry.EntryId FROM ed.ExamInEntry WHERE ExamInEntry.ExamId = {0})", examId);

                    string queryAbits = @"SELECT qAbiturient.Id, qAbiturient.PersonId, E.FacultyId, qAbiturient.EntryId FROM ed.Abiturient AS qAbiturient 
                            INNER JOIN ed.extEntry E ON E.Id = qAbiturient.EntryId
                            LEFT JOIN ed.Person ON qAbiturient.PersonId = Person.Id
                            LEFT JOIN ed.extProtocol ON extProtocol.AbiturientId = qAbiturient.Id WHERE 1 = 1 ";

                    DataSet ds = bdc.GetDataSet(queryAbits + GetAbitFilterString(iFacultyId) + flt_backDoc + flt_enable + flt_protocol + flt_status + flt_mark + flt_hasExam /*+ flt_hasEge*/, new SortedList<string, object>() { { "@ExamId", examId } });

                    var Fac = context.SP_Faculty.Where(x => x.Id == iFacultyId).Select(x => x.Name).FirstOrDefault();
                    var Ex = context.Exam.Where(x => x.Id == examId).Select(x => x.ExamName.Name).FirstOrDefault();

                    wtc.ZeroCount();
                    wtc.SetMax(ds.Tables[0].Rows.Count);
                    wtc.SetText("Зачтено оценок: " + marksCount + "; " + Fac + "/" + Ex);

                    try
                    {
                        foreach (DataRow dsRow in ds.Tables[0].Rows)
                        {
                            wtc.PerformStep();

                            int? balls = null;
                            Guid abId = new Guid(dsRow["Id"].ToString());
                            Guid persId = new Guid(dsRow["PersonId"].ToString());
                            Guid entryId = new Guid(dsRow["EntryId"].ToString());

                            int? exInEntryId = (from eie in context.extExamInEntry
                                                where eie.EntryId == entryId && eie.ExamId == examId
                                                select eie.Id).FirstOrDefault();

                            if (exInEntryId == null)
                                continue;

                            Guid egeCertificateId = Guid.Empty;

                            if (examId != examInostr)
                            {
                                var lBalls =
                                    (from emm in context.hlpEgeMarkMaxApproved
                                     join ete in context.EgeToExam on emm.EgeExamNameId equals ete.EgeExamNameId
                                     join em in context.EgeMark on emm.EgeMarkId equals em.Id
                                     //join ec in context.EgeCertificate on emm.EgeCertificateId equals ec.Id
                                     where emm.PersonId == persId && ete.ExamId == examId
                                     //&& (ec.FBSStatusId == 1 || ec.FBSStatusId == 4)
                                     select new
                                     {
                                         em.Value,
                                         em.EgeCertificateId
                                     }).ToList();
                                if (lBalls.Count() == 0)
                                    continue;
                                balls = lBalls.OrderByDescending(x => x.Value).FirstOrDefault().Value;
                                egeCertificateId = lBalls.OrderByDescending(x => x.Value).FirstOrDefault().EgeCertificateId;
                            }
                            else
                            {
                                List<int> lstInostr = (from ete in context.EgeToExam
                                                       where ete.ExamId == examInostr
                                                       select ete.EgeExamNameId).ToList<int>();

                                if (dsRow["FacultyId"].ToString() == filFacId.ToString())
                                {
                                    int? egeExamNameId = (from etl in context.EgeToLanguage
                                                          join ab in context.qAbiturient
                                                          on etl.LanguageId equals ab.LanguageId
                                                          where etl.ExamId == examInostr && ab.Id == abId
                                                          select etl.EgeExamNameId).FirstOrDefault();

                                    if (egeExamNameId != null)
                                    {
                                        var lBalls =
                                            (from emm in context.extEgeMarkMaxAbitApproved
                                             where emm.AbiturientId == abId && emm.EgeExamNameId == egeExamNameId
                                             select new
                                             {
                                                 emm.Value,
                                                 emm.EgeCertificateId
                                             }).ToList();
                                        if (lBalls.Count() == 0)
                                            continue;
                                        balls = lBalls.OrderByDescending(x => x.Value).FirstOrDefault().Value;
                                        egeCertificateId = lBalls.OrderByDescending(x => x.Value).FirstOrDefault().EgeCertificateId;
                                    }
                                }
                                else
                                {
                                    //int cntEM = (from emm in context.extEgeMarkMaxAbitApproved
                                    //             where lstInostr.Contains(emm.EgeExamNameId) && emm.AbiturientId == abId
                                    //             select emm.EgeMarkId).Count();

                                    //if (cntEM > 1)
                                    //{
                                    //    int? egeExamNameId = (from etl in context.EgeToLanguage
                                    //                          join ab in context.qAbiturient
                                    //                          on etl.LanguageId equals ab.LanguageId
                                    //                          where etl.ExamId == examInostr && ab.Id == abId
                                    //                          select etl.EgeExamNameId).FirstOrDefault();

                                    //    if (egeExamNameId != null)
                                    //    {
                                    //        var lBalls =
                                    //            (from emm in context.extEgeMarkMaxAbitApproved
                                    //             where emm.AbiturientId == abId && emm.EgeExamNameId == egeExamNameId
                                    //             select new
                                    //             {
                                    //                 emm.Value,
                                    //                 emm.EgeCertificateId
                                    //             }).ToList();
                                    //        if (lBalls.Count() == 0)
                                    //            continue;
                                    //        balls = lBalls.OrderByDescending(x => x.Value).FirstOrDefault().Value;
                                    //        egeCertificateId = lBalls.OrderByDescending(x => x.Value).FirstOrDefault().EgeCertificateId;
                                    //    }
                                    //}
                                    //else
                                    //{

                                    var lBalls =
                                    //    (from emm in context.extEgeMarkMaxAbitApproved
                                    //     join ete in context.EgeToExam on emm.EgeExamNameId equals ete.EgeExamNameId
                                    //     join ec in context.EgeCertificate on emm.EgeCertificateId equals ec.Id
                                    //     where emm.AbiturientId == abId && ete.ExamId == examId
                                    //     && (ec.FBSStatusId == 1 || ec.FBSStatusId == 4)
                                    //     select new
                                    //     {
                                    //         emm.Value,
                                    //         emm.EgeCertificateId
                                    //     }).ToList();
                                    (from emm in context.hlpEgeMarkMaxApproved
                                     join ete in context.EgeToExam on emm.EgeExamNameId equals ete.EgeExamNameId
                                     join em in context.EgeMark on emm.EgeMarkId equals em.Id
                                     //join ec in context.EgeCertificate on emm.EgeCertificateId equals ec.Id
                                     where emm.PersonId == persId && ete.ExamId == examId
                                     //&& (ec.FBSStatusId == 1 || ec.FBSStatusId == 4)
                                     select new
                                     {
                                         em.Value,
                                         em.EgeCertificateId
                                     }).ToList();
                                    if (lBalls.Count() == 0)
                                        continue;
                                    balls = lBalls.OrderByDescending(x => x.Value).FirstOrDefault().Value;
                                    egeCertificateId = lBalls.OrderByDescending(x => x.Value).FirstOrDefault().EgeCertificateId;
                                        
                                    //}
                                }
                            }

                            if (balls != null)
                                context.Mark_Insert(abId, exInEntryId, balls, dtDateExam.Value.Date, true, false, false, null, null, egeCertificateId);
                            else
                                continue;

                            marksCount++;
                        }
                    }
                    catch (Exception exc)
                    {
                        throw new Exception("Ошибка загрузки оценок: " + exc.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка загрузки оценок " + ex.Message);
                wtc.Close();
            }            
        }
    }
}