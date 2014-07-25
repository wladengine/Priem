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
    public partial class FormPriemResults_AbitExamResults : Form
    {
        public FormPriemResults_AbitExamResults()
        {
            InitializeComponent();
        }

        private void btnCalculate_Click(object sender, EventArgs e)
        {
            Calculate();
        }

        private void Calculate()
        {
            using (PriemEntities context = new PriemEntities())
            {
                NewWatch wc = new NewWatch();
                wc.Show();
                wc.SetText("Init...");

                DataTable tblRes = new DataTable();
                tblRes.Columns.Add("FIO", typeof(string));
                tblRes.Columns.Add("FacultyName", typeof(string));
                tblRes.Columns.Add("LP", typeof(string));
                tblRes.Columns.Add("SF", typeof(string));
                tblRes.Columns.Add("SB", typeof(string));
                tblRes.Columns.Add("Exam_Rus", typeof(string));
                tblRes.Columns.Add("Exam_Math", typeof(string));
                tblRes.Columns.Add("Exam_Obsh", typeof(string));
                tblRes.Columns.Add("Exam_Ist", typeof(string));
                tblRes.Columns.Add("Exam_Fiz", typeof(string));
                tblRes.Columns.Add("Exam_Chem", typeof(string));
                tblRes.Columns.Add("Exam_Bio", typeof(string));
                tblRes.Columns.Add("Exam_Inost", typeof(string));
                tblRes.Columns.Add("Exam_Liter", typeof(string));
                tblRes.Columns.Add("Exam_Geogr", typeof(string));
                tblRes.Columns.Add("Exam_Inf", typeof(string));

                tblRes.Columns.Add("TvKonk1_Name", typeof(string));
                tblRes.Columns.Add("TvKonk1_Val", typeof(string));
                tblRes.Columns.Add("TvKonk2_Name", typeof(string));
                tblRes.Columns.Add("TvKonk2_Val", typeof(string));
                tblRes.Columns.Add("TvKonk3_Name", typeof(string));
                tblRes.Columns.Add("TvKonk3_Val", typeof(string));

                tblRes.Columns.Add("MarksSum", typeof(string));
                tblRes.Columns.Add("EntryReason", typeof(string));
                tblRes.Columns.Add("OlympReasonName", typeof(string));
                tblRes.Columns.Add("OlympSubjectName", typeof(string));

                wc.SetText("Получение списка абитуриентов...");
                var AbitTable =
                    (from Abit in context.Abiturient
                     join Faculty in context.SP_Faculty on Abit.Entry.FacultyId equals Faculty.Id
                     join Olymp in context.Olympiads on Abit.Id equals Olymp.AbiturientId
                     join _extEntryView in context.extEntryView on Abit.Id equals _extEntryView.AbiturientId

                     join AbitSum in context.extAbitMarksSum on Abit.Id equals AbitSum.Id into AbitSum2
                     from AbitSum in AbitSum2.DefaultIfEmpty()

                     where Abit.Entry.StudyLevel.LevelGroupId == 1
                     && Olymp.OlympValueId > 4 && Abit.Entry.StudyBasisId == 1 && Abit.Entry.StudyFormId == 1
                     select new
                     {
                         Abit.Id,
                         Abit.EntryId,
                         Abit.PersonId,
                         Abit.Person.Surname,
                         Abit.Person.Name,
                         Abit.Person.SecondName,
                         FacultyName = Faculty.Name,
                         LicenseProgramCode = Abit.Entry.SP_LicenseProgram.Code,
                         LicenseProgramName = Abit.Entry.SP_LicenseProgram.Name,
                         StudyForm = Abit.Entry.StudyForm.Name,
                         StudyBasis = Abit.Entry.StudyBasis.Name,
                         Abit.Entry.StudyBasisId,
                         MarksSum = AbitSum.TotalSum,
                         Abit.CompetitionId
                     }).Distinct();

                int iCnt = 0;
                int iMax = AbitTable.Count();
                wc.SetMax(iMax);
                foreach (var AbEnt in AbitTable)
                {
                    wc.SetText("Построение оценок... " + (iCnt++).ToString() + " из " + iMax.ToString());
                    wc.PerformStep();
                    DataRow rw = tblRes.NewRow();
                    rw["FIO"] = AbEnt.Surname + " " + AbEnt.Name + " " + AbEnt.SecondName;
                    rw["FacultyName"] = AbEnt.FacultyName;
                    rw["LP"] = AbEnt.LicenseProgramCode + " " + AbEnt.LicenseProgramName;
                    rw["SF"] = AbEnt.StudyForm;
                    rw["SB"] = AbEnt.StudyBasis;

                //    tblRes.Columns.Add("", typeof(int?));
                //tblRes.Columns.Add("", typeof(int?));
                //tblRes.Columns.Add("", typeof(int?));
                //tblRes.Columns.Add("", typeof(int?));
                //tblRes.Columns.Add("", typeof(int?));
                //tblRes.Columns.Add("", typeof(int?));
                //tblRes.Columns.Add("", typeof(int?));
                //tblRes.Columns.Add("", typeof(int?));
                //tblRes.Columns.Add("", typeof(int?));
                //tblRes.Columns.Add("", typeof(int?));
                //tblRes.Columns.Add("", typeof(int?));

                    var lstExams = context.ExamInEntry.Where(x => x.EntryId == AbEnt.EntryId).Select(x => x.ExamId);

                    int iExam_Rus = context.EgeMark.Where(x => x.EgeCertificate.PersonId == AbEnt.PersonId && (x.EgeCertificate.FBSStatusId == 1 || x.EgeCertificate.FBSStatusId == 4) //&& x.EgeExamName.EgeToExam.Where(y => y.Exam.ExamInEntry.Where(z => z.EntryId == AbEnt.EntryId).Count() > 0).Count() > 0
                        && x.EgeExamName.Name == "Русский язык").Select(x => x.Value).DefaultIfEmpty(0).Max();
                    int iExam_Math = context.EgeMark.Where(x => x.EgeCertificate.PersonId == AbEnt.PersonId && (x.EgeCertificate.FBSStatusId == 1 || x.EgeCertificate.FBSStatusId == 4) //&& x.EgeExamName.EgeToExam.Where(y => y.Exam.ExamInEntry.Where(z => z.EntryId == AbEnt.EntryId).Count() > 0).Count() > 0
                        && x.EgeExamName.Name == "Математика").Select(x => x.Value).DefaultIfEmpty(0).Max();
                    int iExam_Obsh = context.EgeMark.Where(x => x.EgeCertificate.PersonId == AbEnt.PersonId && (x.EgeCertificate.FBSStatusId == 1 || x.EgeCertificate.FBSStatusId == 4) //&& x.EgeExamName.EgeToExam.Where(y => y.Exam.ExamInEntry.Where(z => z.EntryId == AbEnt.EntryId).Count() > 0).Count() > 0
                        && x.EgeExamName.Name == "Обществознание").Select(x => x.Value).DefaultIfEmpty(0).Max();
                    int iExam_Ist = context.EgeMark.Where(x => x.EgeCertificate.PersonId == AbEnt.PersonId && (x.EgeCertificate.FBSStatusId == 1 || x.EgeCertificate.FBSStatusId == 4) //&& x.EgeExamName.EgeToExam.Where(y => y.Exam.ExamInEntry.Where(z => z.EntryId == AbEnt.EntryId).Count() > 0).Count() > 0
                        && x.EgeExamName.Name == "История").Select(x => x.Value).DefaultIfEmpty(0).Max();
                    int iExam_Fiz = context.EgeMark.Where(x => x.EgeCertificate.PersonId == AbEnt.PersonId && (x.EgeCertificate.FBSStatusId == 1 || x.EgeCertificate.FBSStatusId == 4) //&& x.EgeExamName.EgeToExam.Where(y => y.Exam.ExamInEntry.Where(z => z.EntryId == AbEnt.EntryId).Count() > 0).Count() > 0
                        && x.EgeExamName.Name == "Физика").Select(x => x.Value).DefaultIfEmpty(0).Max();
                    int iExam_Chem = context.EgeMark.Where(x => x.EgeCertificate.PersonId == AbEnt.PersonId && (x.EgeCertificate.FBSStatusId == 1 || x.EgeCertificate.FBSStatusId == 4) //&& x.EgeExamName.EgeToExam.Where(y => y.Exam.ExamInEntry.Where(z => z.EntryId == AbEnt.EntryId).Count() > 0).Count() > 0
                        && x.EgeExamName.Name == "Химия").Select(x => x.Value).DefaultIfEmpty(0).Max();
                    int iExam_Bio = context.EgeMark.Where(x => x.EgeCertificate.PersonId == AbEnt.PersonId && (x.EgeCertificate.FBSStatusId == 1 || x.EgeCertificate.FBSStatusId == 4) //&& x.EgeExamName.EgeToExam.Where(y => y.Exam.ExamInEntry.Where(z => z.EntryId == AbEnt.EntryId).Count() > 0).Count() > 0
                        && x.EgeExamName.Name == "Биология").Select(x => x.Value).DefaultIfEmpty(0).Max();
                    int iExam_Inost = context.EgeMark.Where(x => x.EgeCertificate.PersonId == AbEnt.PersonId && (x.EgeCertificate.FBSStatusId == 1 || x.EgeCertificate.FBSStatusId == 4) //&& x.EgeExamName.EgeToExam.Where(y => y.Exam.ExamInEntry.Where(z => z.EntryId == AbEnt.EntryId).Count() > 0).Count() > 0
                        && (x.EgeExamName.Name == "Английский язык" || x.EgeExamName.Name == "Немецкий язык" || x.EgeExamName.Name == "Французский язык" || x.EgeExamName.Name == "Испанский язык")).Select(x => x.Value).DefaultIfEmpty(0).Max();
                    int iExam_Liter = context.EgeMark.Where(x => x.EgeCertificate.PersonId == AbEnt.PersonId && (x.EgeCertificate.FBSStatusId == 1 || x.EgeCertificate.FBSStatusId == 4) //&& x.EgeExamName.EgeToExam.Where(y => y.Exam.ExamInEntry.Where(z => z.EntryId == AbEnt.EntryId).Count() > 0).Count() > 0
                        && x.EgeExamName.Name == "Литература").Select(x => x.Value).DefaultIfEmpty(0).Max();
                    int iExam_Geogr = context.EgeMark.Where(x => x.EgeCertificate.PersonId == AbEnt.PersonId && (x.EgeCertificate.FBSStatusId == 1 || x.EgeCertificate.FBSStatusId == 4) //&& x.EgeExamName.EgeToExam.Where(y => y.Exam.ExamInEntry.Where(z => z.EntryId == AbEnt.EntryId).Count() > 0).Count() > 0
                        && x.EgeExamName.Name == "География").Select(x => x.Value).DefaultIfEmpty(0).Max();
                    int iExam_Inf = context.EgeMark.Where(x => x.EgeCertificate.PersonId == AbEnt.PersonId && (x.EgeCertificate.FBSStatusId == 1 || x.EgeCertificate.FBSStatusId == 4) //&& x.EgeExamName.EgeToExam.Where(y => y.Exam.ExamInEntry.Where(z => z.EntryId == AbEnt.EntryId).Count() > 0).Count() > 0
                        && x.EgeExamName.Name == "Информатика и ИКТ").Select(x => x.Value).DefaultIfEmpty(0).Max();
                    rw["Exam_Rus"] = iExam_Rus == 0 ? "" : iExam_Rus.ToString();
                        //context.Mark.Where(x => x.AbiturientId == AbEnt.Id && x.ExamInEntry.Exam.ExamName.Name == "Русский язык")
                        //.Select(x => new { Value = (int)x.Value, x.IsFromOlymp, x.ExamVedId }).ToList()
                        //.Select(x => x.Value + (x.IsFromOlymp ? " (О)" : (x.ExamVedId.HasValue ? " (Э)" : ""))).FirstOrDefault();
                    rw["Exam_Math"] = iExam_Math == 0 ? "" : iExam_Math.ToString();
                        //context.Mark.Where(x => x.AbiturientId == AbEnt.Id && x.ExamInEntry.Exam.ExamName.Name == "Математика")
                        //.Select(x => new { Value = (int)x.Value, x.IsFromOlymp, x.ExamVedId }).ToList()
                        //.Select(x => x.Value + (x.IsFromOlymp ? " (О)" : (x.ExamVedId.HasValue ? " (Э)" : ""))).FirstOrDefault();
                    rw["Exam_Obsh"] = iExam_Obsh == 0 ? "" : iExam_Obsh.ToString();
                        //context.Mark.Where(x => x.AbiturientId == AbEnt.Id && x.ExamInEntry.Exam.ExamName.Name == "Обществознание")
                        //.Select(x => new { Value = (int)x.Value, x.IsFromOlymp, x.ExamVedId }).ToList()
                        //.Select(x => x.Value + (x.IsFromOlymp ? " (О)" : (x.ExamVedId.HasValue ? " (Э)" : ""))).FirstOrDefault();
                    rw["Exam_Ist"] = iExam_Ist == 0 ? "" : iExam_Ist.ToString();
                        //context.Mark.Where(x => x.AbiturientId == AbEnt.Id && x.ExamInEntry.Exam.ExamName.Name == "История")
                        //.Select(x => new { Value = (int)x.Value, x.IsFromOlymp, x.ExamVedId }).ToList()
                        //.Select(x => x.Value + (x.IsFromOlymp ? " (О)" : (x.ExamVedId.HasValue ? " (Э)" : ""))).FirstOrDefault();
                    rw["Exam_Fiz"] = iExam_Fiz == 0 ? "" : iExam_Fiz.ToString();
                        //context.Mark.Where(x => x.AbiturientId == AbEnt.Id && x.ExamInEntry.Exam.ExamName.Name == "Физика")
                        //.Select(x => new { Value = (int)x.Value, x.IsFromOlymp, x.ExamVedId }).ToList()
                        //.Select(x => x.Value + (x.IsFromOlymp ? " (О)" : (x.ExamVedId.HasValue ? " (Э)" : ""))).FirstOrDefault();
                    rw["Exam_Chem"] = iExam_Chem == 0 ? "" : iExam_Chem.ToString();
                        //context.Mark.Where(x => x.AbiturientId == AbEnt.Id && x.ExamInEntry.Exam.ExamName.Name == "Химия")
                        //.Select(x => new { Value = (int)x.Value, x.IsFromOlymp, x.ExamVedId }).ToList()
                        //.Select(x => x.Value + (x.IsFromOlymp ? " (О)" : (x.ExamVedId.HasValue ? " (Э)" : ""))).FirstOrDefault();
                    rw["Exam_Bio"] = iExam_Bio == 0 ? "" : iExam_Bio.ToString();
                        //context.Mark.Where(x => x.AbiturientId == AbEnt.Id && x.ExamInEntry.Exam.ExamName.Name == "Биология")
                        //.Select(x => new { Value = (int)x.Value, x.IsFromOlymp, x.ExamVedId }).ToList()
                        //.Select(x => x.Value + (x.IsFromOlymp ? " (О)" : (x.ExamVedId.HasValue ? " (Э)" : ""))).FirstOrDefault();
                    rw["Exam_Inost"] = iExam_Inost == 0 ? "" : iExam_Inost.ToString();
                        //context.Mark.Where(x => x.AbiturientId == AbEnt.Id && x.ExamInEntry.Exam.ExamName.Name == "Иностранный язык")
                        //.Select(x => new { Value = (int)x.Value, x.IsFromOlymp, x.ExamVedId }).ToList()
                        //.Select(x => x.Value + (x.IsFromOlymp ? " (О)" : (x.ExamVedId.HasValue ? " (Э)" : ""))).FirstOrDefault();
                    rw["Exam_Liter"] = iExam_Liter == 0 ? "" : iExam_Liter.ToString();
                        //context.Mark.Where(x => x.AbiturientId == AbEnt.Id && x.ExamInEntry.Exam.ExamName.Name == "Литература")
                        //.Select(x => new { Value = (int)x.Value, x.IsFromOlymp, x.ExamVedId }).ToList()
                        //.Select(x => x.Value + (x.IsFromOlymp ? " (О)" : (x.ExamVedId.HasValue ? " (Э)" : ""))).FirstOrDefault();
                    rw["Exam_Geogr"] = iExam_Geogr == 0 ? "" : iExam_Geogr.ToString();
                        //context.Mark.Where(x => x.AbiturientId == AbEnt.Id && x.ExamInEntry.Exam.ExamName.Name == "География")
                        //.Select(x => new { Value = (int)x.Value, x.IsFromOlymp, x.ExamVedId }).ToList()
                        //.Select(x => x.Value + (x.IsFromOlymp ? " (О)" : (x.ExamVedId.HasValue ? " (Э)" : ""))).FirstOrDefault();
                    rw["Exam_Inf"] = iExam_Inf == 0 ? "" : iExam_Inf.ToString();
                        //context.Mark.Where(x => x.AbiturientId == AbEnt.Id && x.ExamInEntry.Exam.ExamName.Name == "Информатика и информационно-коммуникационные технологии (ИКТ)")
                        //.Select(x => new { Value = (int)x.Value, x.IsFromOlymp, x.ExamVedId }).ToList()
                        //.Select(x => x.Value + (x.IsFromOlymp ? " (О)" : (x.ExamVedId.HasValue ? " (Э)" : ""))).FirstOrDefault();

                    rw["MarksSum"] = AbEnt.MarksSum.HasValue ? AbEnt.MarksSum.Value.ToString() : "";

                    if (AbEnt.CompetitionId == 1)//б/э - бывают победители-призёры всероссов, а бывают другие
                    {
                        var vseross = context.extOlympiads.Where(x => x.AbiturientId == AbEnt.Id && x.OlympTypeId == 2 && x.OlympValueId > 4);
                        if (vseross.Count() > 0)
                            rw["EntryReason"] = "2";
                        else
                        {
                            rw["EntryReason"] = "3";
                            //var olymps = context.Olympiads.Where(x => x.AbiturientId == AbEnt.Id && x.OlympTypeId != 2 && x.OlympValueId > 4).OrderBy(x => x.OlympValue.SortOrder);
                            //if (olymps.Count() > 0)
                            //    rw["OlympReasonName"] = context.OlympName.Where(x => x.Id == olymps.FirstOrDefault().OlympNameId).Select(x => x.Name).FirstOrDefault();
                        }
                    }
                    else if (AbEnt.CompetitionId == 2)//в/к
                        rw["EntryReason"] = "5";
                    else if (AbEnt.CompetitionId == 6)//целевики
                        rw["EntryReason"] = "4";
                    //потом этот кусок удалить
                    var olymps = context.Olympiads.Where(x => x.AbiturientId == AbEnt.Id && x.OlympValueId > 4).OrderBy(x => x.OlympValue.SortOrder);
                    if (olymps.Count() > 0)
                    {
                        rw["OlympReasonName"] =
                            olymps.FirstOrDefault().OlympTypeId == 2 ? "Всеросс" :
                            context.OlympName.Where(x => x.Id == olymps.FirstOrDefault().OlympNameId).Select(x => x.Name).FirstOrDefault();
                        rw["OlympSubjectName"] = olymps.FirstOrDefault().OlympSubjectId.HasValue ? olymps.FirstOrDefault().OlympSubject.Name : "";
                    }
                    //до сюда удалять
                    var TvConk = context.extExamInEntry.Where(x => x.ExamName.Contains("творческий конкурс") && x.EntryId == AbEnt.EntryId);
                    if (TvConk.Count() > 0)
                    {
                        int i = 1;
                        foreach (var conk in TvConk.OrderBy(x => x.ExamNameId))
                        {
                            rw["TvKonk" + i.ToString() + "_Name"] = conk.ExamName;
                            rw["TvKonk" + i.ToString() + "_Val"] = context.Mark.Where(x => x.AbiturientId == AbEnt.Id && x.ExamInEntryId == conk.Id).Select(x => x.Value).FirstOrDefault().ToString();
                            i++;
                        }
                    }

                    tblRes.Rows.Add(rw);
                }

                dgvAbitList.DataSource = tblRes;
                wc.Close();
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "CSV files| *.csv";
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                using (System.IO.StreamWriter sw = new System.IO.StreamWriter(sfd.FileName, false, Encoding.GetEncoding(1251)))
                {
                    foreach (DataGridViewRow rw in dgvAbitList.Rows)
                    {
                        string rwdata = "";
                        foreach (DataGridViewCell c in rw.Cells)
                            rwdata += c.Value.ToString() + ";";

                        sw.WriteLine(rwdata);
                    }

                    sw.Flush();
                }
            }
        }
    }
}
