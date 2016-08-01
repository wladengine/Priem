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
//            string quer = @"SELECT DISTINCT extEntry.FacultyName AS [Факультет], 
//COUNT(DISTINCT t.AbiturientId) AS [Абитуриентов], 
//(
//	SELECT CONVERT(nvarchar, MAX(extEnableProtocol.Date), 104) 
//	+ ' ' + CONVERT(nvarchar, MAX(extEnableProtocol.Date), 108)
//	FROM ed.extEnableProtocol WHERE extEnableProtocol.StudyLevelGroupId = 1 
//	AND extEnableProtocol.FacultyId = extEntry.FacultyId
//)  AS [Дата последнего протокола о допуске]
//FROM
//(
//    SELECT 
//		CASE WHEN Mark.Id IS NULL AND Person_AdditionalInfo.EgeInSPbgu = 0
//		THEN Abiturient.Id 
//		ELSE NULL 
//		END
//		AS AbiturientId, Abiturient.EntryId
//    FROM ed.Abiturient
//    INNER JOIN ed.Person_AdditionalInfo ON Person_AdditionalInfo.PersonId = Abiturient.PersonId
//    INNER JOIN ed.extExamInEntry ON extExamInEntry.EntryId = Abiturient.EntryId
//    INNER JOIN ed.EgeToExam ON EgeToExam.ExamId = extExamInEntry.ExamId
//    INNER JOIN ed.extEnableProtocol ON extEnableProtocol.AbiturientId = Abiturient.Id
//    LEFT JOIN ed.Mark ON Mark.ExamInEntryBlockUnitId = extExamInEntry.Id AND Mark.AbiturientId = Abiturient.Id
//    WHERE extEnableProtocol.IsOld = 0 AND extEnableProtocol.Excluded = 0
//    AND Abiturient.BackDoc = 0 AND Abiturient.NotEnabled = 0
//    AND Abiturient.CompetitionId NOT IN (1, 8)
//) t
//INNER JOIN ed.extEntry ON extEntry.Id = t.EntryId
//WHERE extEntry.StudyLevelGroupId = 1
//GROUP BY extEntry.FacultyName, extEntry.FacultyId
//ORDER BY 1";

            BackgroundWorker bw = new BackgroundWorker();
            bw.DoWork += (sender, e) =>
            {
                string quer = @"SELECT DISTINCT extEntry.FacultyName AS [Факультет], 
COUNT(DISTINCT t.AbiturientId) AS [Абитуриентов], 
(
	SELECT CONVERT(nvarchar, MAX(extEnableProtocol.Date), 104) 
	+ ' ' + CONVERT(nvarchar, MAX(extEnableProtocol.Date), 108)
	FROM ed.extEnableProtocol WHERE extEnableProtocol.StudyLevelGroupId = 1 
	AND extEnableProtocol.FacultyId = extEntry.FacultyId
)  AS [Дата последнего протокола о допуске]
FROM
(
    SELECT 
		CASE WHEN Mark.Id IS NULL AND Person_AdditionalInfo.EgeInSPbgu = 0
		THEN Abiturient.Id 
		ELSE NULL 
		END
		AS AbiturientId, Abiturient.EntryId
    FROM ed.Abiturient
    INNER JOIN ed.Person_AdditionalInfo ON Person_AdditionalInfo.PersonId = Abiturient.PersonId
    INNER JOIN ed.extExamInEntry ON extExamInEntry.EntryId = Abiturient.EntryId
    INNER JOIN ed.EgeToExam ON EgeToExam.ExamId = extExamInEntry.ExamId
    INNER JOIN ed.extEnableProtocol ON extEnableProtocol.AbiturientId = Abiturient.Id
    LEFT JOIN ed.Mark ON Mark.ExamInEntryBlockUnitId = extExamInEntry.Id AND Mark.AbiturientId = Abiturient.Id
    WHERE extEnableProtocol.IsOld = 0 AND extEnableProtocol.Excluded = 0
    AND Abiturient.BackDoc = 0 AND Abiturient.NotEnabled = 0
    AND Abiturient.CompetitionId NOT IN (1, 8)
) t
INNER JOIN ed.extEntry ON extEntry.Id = t.EntryId
WHERE extEntry.StudyLevelGroupId = 1
GROUP BY extEntry.FacultyName, extEntry.FacultyId
ORDER BY 1";
                e.Result = MainClass.Bdc.GetDataSet(quer).Tables[0];
            };
            bw.RunWorkerCompleted += (sender, e) =>
            {
                gbLoading.Visible = false;
                if (!e.Cancelled)
                {
                    dgvProtocols.DataSource = e.Result;
                    dgvProtocols.Columns["Факультет"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
                    dgvProtocols.Columns["Абитуриентов"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
                    dgvProtocols.Columns["Дата последнего протокола о допуске"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
                }
            };
            gbLoading.Visible = true;
            bw.RunWorkerAsync();
            //dgvProtocols.DataSource = MainClass.Bdc.GetDataSet(quer).Tables[0];
            //dgvProtocols.Columns["Факультет"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
            //dgvProtocols.Columns["Абитуриентов"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
            //dgvProtocols.Columns["Дата последнего протокола о допуске"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
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

            bool deepScan = chbDeepScan.Checked;

            try
            {
                List<int> lstFac = new List<int>();
                if (FacultyId == null)
                {
                    using (PriemEntities context = new PriemEntities())
                    {
                        lstFac = context.extEntry.Where(x => x.StudyLevelGroupId == 1).Select(x => x.FacultyId).Distinct().ToList();
                    }
                }
                else
                    lstFac.Add(FacultyId.Value);

                List<KeyValuePair<int, int>> lstPairs = new List<KeyValuePair<int, int>>();
                foreach (int iFacId in lstFac)
                {
                    if (ExamId == null)
                    {
                        using (PriemEntities context = new PriemEntities())
                        {
                            var ent = Exams.GetExamsWithFilters(context, MainClass.lstStudyLevelGroupId, iFacId, null, null, null, null, null, null, null, null)
                                    .Where(c => !c.IsAdditional && !c.IsSecond && !c.IsGosLine);
                            foreach (var exInEnt in ent.Select(x => x.ExamId).Distinct())
                                lstPairs.Add(new KeyValuePair<int, int>(iFacId, exInEnt));
                        }
                    }
                    else
                        lstPairs.Add(new KeyValuePair<int, int>(iFacId, ExamId.Value));
                }

                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += (sender, e) =>
                {
                    wtc = new NewWatch(2);
                    wtc.Show();
                    wtc.TopMost = true;
                    marksCount = 0;

                    foreach (KeyValuePair<int, int> kvp in lstPairs)
                        SetMarksForExam(kvp.Value, kvp.Key, deepScan);

                    wtc.Close();
                };
                bw.RunWorkerCompleted += (sender, e) =>
                {
                    
                    btnOk.Enabled = true;
                    UpdateGridAbits();
                    MessageBox.Show(string.Format("Зачтено {0} оценок", marksCount));

                    MainClass._1k_LastEgeMarkLoadTime = DateTime.Now;
                };

                bw.RunWorkerAsync();
                btnOk.Enabled = false;
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка загрузки оценок " + ex.Message);
            }
        }

        private void SetMarksForExam(int examId, int iFacultyId, bool deepScan)
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
            INNER JOIN ed.extExamInEntry ON Mark.ExamInEntryBlockUnitId = extExamInEntry.Id 
            WHERE Mark.AbiturientId = qAbiturient.Id 
            AND extExamInEntry.ExamId = {0}
            AND extExamInEntry.IsAdditional = 0
    )
    {1}
)", examId, deepScan ? string.Format(@"OR qAbiturient.Id IN
    /*или у абитуриента зачёлся балл ниже, чем есть среди его ЕГЭ*/
    (
        SELECT qMark.AbiturientId
        FROM ed.Mark AS qMark
        INNER JOIN ed.Abiturient ABIT ON ABIT.Id = qMark.AbiturientId
        INNER JOIN ed.hlpEgeMarkMaxApprovedValue ON hlpEgeMarkMaxApprovedValue.PersonId = ABIT.PersonId 
        INNER JOIN ed.extExamInEntry ON extExamInEntry.Id = qMark.ExamInEntryBlockUnitId
        INNER JOIN ed.EgeToExam ON EgeToExam.ExamId = extExamInEntry.ExamId 
        WHERE EgeToExam.EgeExamNameId = hlpEgeMarkMaxApprovedValue.EgeExamNameId 
        AND qMark.IsFromEge = 1 
        AND qAbiturient.EntryId = ABIT.EntryId
        AND extExamInEntry.ExamId = EgeToExam.ExamId
        AND extExamInEntry.ExamId = {0}
        AND hlpEgeMarkMaxApprovedValue.EgeMarkValue > qMark.Value
    )", examId) : "");
                    //string flt_hasEge = string.Format(" AND Person.Id IN (SELECT PersonId FROM ed.EgeMark LEFT JOIN ed.EgeToExam ON EgeMark.EgeExamNameId = EgeToExam.EgeExamNameId WHERE EgeToExam.ExamId = @ExamId)", examId);
                    string flt_hasExam = string.Format(" AND qAbiturient.EntryId IN (SELECT extExamInEntry.EntryId FROM ed.extExamInEntry WHERE extExamInEntry.ExamId = {0})", examId);

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

                            Guid? exInEntryBlockUnitId = (from eie in context.extExamInEntry
                                                where eie.EntryId == entryId && eie.ExamId == examId
                                                select eie.Id).FirstOrDefault();

                            if (exInEntryBlockUnitId == null)
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
                                context.Mark_Insert(abId, exInEntryBlockUnitId, balls, dtDateExam.Value.Date, true, false, false, null, null, egeCertificateId);
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