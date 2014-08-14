using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using EducServLib;
using System.Xml;

namespace Priem
{
    public partial class StatFormGSGUForm2 : Form
    {
        public StatFormGSGUForm2()
        {
            InitializeComponent();
            this.MdiParent = MainClass.mainform;
            FillCombos();
        }

        private int? StudyLevelId
        {
            get { return ComboServ.GetComboIdInt(cbStudyLevel); }
        }

        private void FillCombos()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.StudyLevel.Where(x => x.LevelGroupId == 1 || x.LevelGroupId == 2).Select(x => new { x.Id, x.Name }).ToList().Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();

                ComboServ.FillCombo(cbStudyLevel, src, false, true);
            }
        }

        private void btnStartImport_Click(object sender, EventArgs e)
        {
            XmlDocument doc = new XmlDocument();
            var declaration = doc.CreateXmlDeclaration("1.0", "utf-8", "");
            doc.AppendChild(declaration);
            var rootNode = doc.AppendChild(doc.CreateNode(XmlNodeType.Element, "root", ""));
            XmlAttribute attr = doc.CreateAttribute("id");
            attr.Value = "2609";
            rootNode.Attributes.Append(attr);
            using (PriemEntities context = new PriemEntities())
            {
                int rowNum = 0;
                var ListLP = context.Entry.Where(x => x.KCP.HasValue && x.KCP > 0 && x.StudyLevel.Id == StudyLevelId).Select(x => new { x.StudyFormId, x.StudyBasisId, x.LicenseProgramId, x.SP_LicenseProgram.GSGUCode, x.SP_LicenseProgram.Name, x.SP_LicenseProgram.Code }).Distinct().ToList();
                ProgressForm pf = new ProgressForm();
                pf.SetProgressText("Загрузка данных...");
                pf.MaxPrBarValue = ListLP.Count;
                pf.Show();
                //try
                //{
                    foreach (var LP in ListLP)
                    {
                        pf.PerformStep();
                        if (string.IsNullOrEmpty(LP.GSGUCode))
                        {
                            MessageBox.Show(LP.Code + " " + LP.Name + " - не указан код ГЗГУ");
                            continue;
                        }
                        rowNum++;

                        //номер строки
                        var rwNode = rootNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "lines", ""));
                        attr = doc.CreateAttribute("id");
                        attr.Value = rowNum.ToString();
                        rwNode.Attributes.Append(attr);

                        //ID организации или филиала, предоставляющего данные
                        var node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "oo", ""));
                        node.InnerText = "2609";

                        //ID специальности (по справочнику №2)
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "spec", ""));
                        node.InnerText = LP.GSGUCode;

                        //ID формы обучения (по справочнику №3)
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "fo", ""));
                        node.InnerText = LP.StudyFormId.ToString();

                        //ID формы финансирования (по справочнику №4)
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "ff", ""));
                        node.InnerText = LP.StudyBasisId.ToString();

                        var EV = context.extEntryView.Where(x => x.LicenseProgramId == LP.LicenseProgramId && x.StudyFormId == LP.StudyFormId && x.StudyBasisId == LP.StudyBasisId && x.IsCrimea == 0);

                        // p1-1 Всего зачисленных на места приема граждан
                            int cnt = EV.Count();
                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p1_1", ""));
                            node.InnerText = cnt.ToString();
                        
                        // p2-1 Всего по общему конкурсу
                            int cntCommonComp = EV.Where(x => x.IsCrimea == 0 && x.IsCel == 0 && x.IsQuota == 0 && x.IsBE == 0).Count();
                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p2_1", ""));
                            node.InnerText = cntCommonComp.ToString();


                            var AbitOlympList =
                                (from ev in EV
                                 join mark in context.Mark on ev.AbiturientId equals mark.AbiturientId
                                 where mark.IsFromOlymp == true
                                 && ev.IsCel == 0 && ev.IsQuota == 0 && ev.IsCrimea == 0 && ev.IsBE == 0
                                 select mark.AbiturientId).Distinct();

                        // p2-2 по общему конкурсу, имеющих результаты ЕГЭ 
                            //( без победителей и призеров олимпиад, зачисленных по общему конкурсу с учетом 100 баллов ЕГЭ по олимпиадному предмету
                            int cntCommonCompWithEGE =
                                (from ev in EV
                                 join mark in context.Mark on ev.AbiturientId equals mark.AbiturientId
                                 where mark.IsFromEge == true && !AbitOlympList.Contains(mark.AbiturientId) && ev.IsCel == 0 && ev.IsQuota == 0 && ev.IsCrimea == 0 && ev.IsBE == 0
                                 select mark.AbiturientId).Distinct().Count();

                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p2_2", ""));
                            node.InnerText = (cntCommonCompWithEGE).ToString();

                        // p2-3 Победители и призеры олимпиад, зачисленные по общему конкурсу с учетом 100 баллов ЕГЭ по олимпиадному предмету
                            int cntOlympLike100Balls =
                                (from exEV in context.extEntryView
                                 join mrk in context.Mark on exEV.AbiturientId equals mrk.AbiturientId
                                 where mrk.IsFromOlymp
                                 && exEV.LicenseProgramId == LP.LicenseProgramId
                                 && exEV.StudyFormId == LP.StudyFormId
                                 && exEV.StudyBasisId == LP.StudyBasisId
                                 && exEV.IsCrimea == 0 && exEV.IsCel == 0 && exEV.IsQuota == 0 && exEV.IsBE == 0
                                 select exEV.AbiturientId).Distinct().Count();
                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p2_3", ""));
                            node.InnerText = cntOlympLike100Balls.ToString();

                            var AbitEgeList =
                                    (from ev in EV
                                     join mark in context.Mark on ev.AbiturientId equals mark.AbiturientId
                                     where mark.IsFromEge == true
                                     && ev.IsCel == 0 && ev.IsQuota == 0 && ev.IsCrimea == 0 && ev.IsBE == 0
                                     select mark.AbiturientId).Distinct();

                        // p2-4 по общему конкурсу, не имеющих результаты ЕГЭ (абитуриенты, сдававшие вступительные испытания, форма которых определяется вузом самостоятельно)
                            int cntCommonCompWithoutEGE =
                                (from ev in EV
                                 join mark in context.Mark on ev.AbiturientId equals mark.AbiturientId
                                 where !AbitEgeList.Contains(mark.AbiturientId) && ev.IsCel == 0 && ev.IsQuota == 0 && ev.IsCrimea == 0 && ev.IsBE == 0
                                 select mark.AbiturientId).Distinct().Count();

                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p2_4", ""));
                            node.InnerText = (cntCommonCompWithoutEGE).ToString();
                            

                        //p2-5 по общему конкурсу, имеющих результаты ЕГЭ и сдававших вступительные испытания творческой и (или) профессиональной направленности ( без победителей 
                            // и призеров олимпиад, зачисленных по общему конкурсу с учетом 100 баллов егэ по по олимпиадному предмету ( из р2_1)
                            int cntCommonCompWithEGEandExamIsAdd =
                                (from ev in EV
                                 join mark in context.Mark on ev.AbiturientId equals mark.AbiturientId
                                 join ent in context.ExamInEntry on ev.EntryId equals ent.EntryId
                                 join exam in context.Exam on ent.ExamId equals exam.Id
                                 where mark.IsFromEge == true && !AbitOlympList.Contains(mark.AbiturientId)
                                 && ev.IsQuota == 0 && ev.IsCrimea == 0 && ev.IsBE == 0 && ev.IsCel == 0 
                                 && exam.IsAdditional == true
                                 select mark.AbiturientId).Distinct().Count();

                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p2_5", ""));
                            node.InnerText = (cntCommonCompWithEGEandExamIsAdd).ToString();   

                        // р3-1 в пределах квоты целевого приема лица, имеющие результаты ЕГЭ
                            int cntCelCompWithEGE =
                                (from ev in EV
                                 join mark in context.Mark on ev.AbiturientId equals mark.AbiturientId
                                 where mark.IsFromEge == true && ev.IsCel == 1 && ev.IsQuota == 0 && ev.IsCrimea == 0 && ev.IsBE == 0
                                 select mark.AbiturientId).Distinct().Count();

                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p3_1", ""));
                            node.InnerText = (cntCelCompWithEGE).ToString();  
                        //р3-2 в пределах квоты целевого приема лица, не имеющих результатов ЕГЭ
                            int cntCelCompWithoutEGE =
                                (from ev in EV
                                 join mark in context.Mark on ev.AbiturientId equals mark.AbiturientId
                                 where !AbitEgeList.Contains(mark.AbiturientId) && ev.IsCel == 1 && ev.IsQuota == 0 && ev.IsCrimea == 0 && ev.IsBE == 0
                                 select mark.AbiturientId).Distinct().Count();

                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p3_2", ""));
                            node.InnerText = (cntCommonCompWithoutEGE).ToString();
                        // р4-1 в пределах квоты приема, имеющих особое право и имеющие результаты ЕГЭ
                            int cntQuotaCompWithEGE =
                                (from ev in EV
                                 join mark in context.Mark on ev.AbiturientId equals mark.AbiturientId
                                 where mark.IsFromEge == true && ev.IsCel == 0 && ev.IsQuota == 1 && ev.IsCrimea == 0 && ev.IsBE == 0
                                 select mark.AbiturientId).Distinct().Count();

                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p4_1", ""));
                            node.InnerText = (cntQuotaCompWithEGE).ToString();  
                        // р4-2 п пределах квоты приема, имеющих особое право и лица, имеющие особое право, зачисленные вне конкурса, не имеющие результатов ЕГЭ
                            int cntQuotaCompWithoutEGE =
                                (from ev in EV
                                 join mark in context.Mark on ev.AbiturientId equals mark.AbiturientId
                                 where !AbitEgeList.Contains(mark.AbiturientId)
                                 && ev.IsCel == 0 && ev.IsQuota == 1 && ev.IsCrimea == 0 && ev.IsBE == 0
                                 select mark.AbiturientId).Distinct().Count();

                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p4_2", ""));
                            node.InnerText = (cntQuotaCompWithoutEGE).ToString(); 
                        // Р5 Победители и призеры олимпиад, члены сборных команд Российской Федерации, участвовавшие в международных олимпиадах по 
                            // общеобразовательным предметам и сформированных в порядве, определяемом Минобрнауки РФ, чемпионы и призеры Олимпийских игр,
                            // Параолимпийских игр, Сурдлимпийских игр, зачисленные без вступительных испытаний

                            int OlympCount = EV.Where(x => x.IsBE == 1).Count();
                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p5", ""));
                            node.InnerText = (OlympCount).ToString(); 

                        //р6-1  средний балл зачисленных по общему конкурсу, имеющих результаты ЕГЭ без учета вступительных испытаний творческой и(или) профессиональной
                            // направленности (без победителей и призеров олимпиад, зачисленных по общему конкурсу с учетом 100 баллов ЕГЭ по олимпиадному предмету
                            double AVGMarkEGE =
                                (from ev in EV
                                 join mark in context.Mark on ev.AbiturientId equals mark.AbiturientId
                                 join ent in context.ExamInEntry on ev.EntryId equals ent.EntryId
                                 join exam in context.Exam on ent.ExamId equals exam.Id 
                                 where !AbitOlympList.Contains(ev.AbiturientId)
                                 && ev.IsCel == 0 && ev.IsQuota == 0 && ev.IsCrimea == 0 && ev.IsBE == 0
                                 && exam.IsAdditional == false
                                 select (int)mark.Value).DefaultIfEmpty(0).Average();
                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p6_1", ""));
                            node.InnerText = (AVGMarkEGE).ToString(); 
                        // р6-2 средний балл зачисленных по общему конкурсу, имеющих результаты ЕГЭ и являющихся победителями и призерами олимпиад, зачисленных
                            // с учетом 100 баллов ЕГЭ по олимпиадному предмету, без учета вступительных испытаний творческой и(или) профессиональной направленности
                            double AVGMarkEGEWithOlymp =
                                (from ev in EV
                                 join mark in context.Mark on ev.AbiturientId equals mark.AbiturientId
                                 join ent in context.ExamInEntry on ev.EntryId equals ent.EntryId
                                 join exam in context.Exam on ent.ExamId equals exam.Id
                                 where AbitOlympList.Contains(ev.AbiturientId)
                                 && ev.IsCel == 0 && ev.IsQuota == 0 && ev.IsCrimea == 0 && ev.IsBE == 0
                                 && exam.IsAdditional == false
                                 select (int)mark.Value).DefaultIfEmpty(0).Average();
                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p6_2", ""));
                            node.InnerText = (AVGMarkEGEWithOlymp).ToString(); 

                        // р6-3 средний балл зачисленных по общему конкурсу, имеющих результаты ЕГЭ с учетом вступительных испытаний творческой и(или) профессиональной направленности
                            // без победителей и призеров олимпиад, зачисленных  с учетом 100 баллов ЕГЭ по олимпиадному предмету,
                            double AVGMarkEGEWithAddExam =
                                (from ev in EV
                                 join mark in context.Mark on ev.AbiturientId equals mark.AbiturientId
                                 where ev.IsCel == 0 && ev.IsQuota == 0 && ev.IsCrimea == 0 && ev.IsBE == 0
                                 && !AbitOlympList.Contains(ev.AbiturientId)
                                 select (int)mark.Value).DefaultIfEmpty(0).Average();
                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p6_3", ""));
                            node.InnerText = (AVGMarkEGEWithAddExam).ToString(); 
                        // р6-4 средний балл зачисленных по общему конкурсу, имеющих результаты ЕГЭ с учетом вступительных испытаний творческой и(или) профессиональной направленности
                            // и являющихся победителями и призерами олимпиад, зачисленных  с учетом 100 баллов ЕГЭ по олимпиадному предмету
                            double AVGMarkEGEWithOlympWithAddExam =
                                (from ev in EV
                                 join mark in context.Mark on ev.AbiturientId equals mark.AbiturientId
                                 where ev.IsCel == 0 && ev.IsQuota == 0 && ev.IsCrimea == 0 && ev.IsBE == 0
                                 && AbitOlympList.Contains(ev.AbiturientId)
                                 select (int)mark.Value).DefaultIfEmpty(0).Average();
                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p6_4", ""));
                            node.InnerText = (AVGMarkEGEWithOlympWithAddExam).ToString(); 

                        //р7-1 Средний балл в пределах квоты целевого приема, имеющих результаты ЕГЭ без учета вступительных испытаний творческой и (или) профессиональной
                            // направленности
                            double AVGMarkEGE_Cel =
                                (from ev in EV
                                 join mark in context.Mark on ev.AbiturientId equals mark.AbiturientId
                                 join ent in context.ExamInEntry on ev.EntryId equals ent.EntryId
                                 join exam in context.Exam on ent.ExamId equals exam.Id 
                                 where ev.IsCel == 1 && ev.IsQuota == 0 && ev.IsCrimea == 0 && ev.IsBE == 0
                                 && exam.IsAdditional == false
                                 select (int)mark.Value).DefaultIfEmpty(0).Average();
                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p7_1", ""));
                            node.InnerText = (AVGMarkEGE_Cel).ToString(); 
                        // р7-2 Средний балл в пределах квоты лиц, имеющих особое право, имеющих результаты ЕГЭ без учета
                            // вступительных испытаний испытаний творческой и (или) профессиональной  направленности
                            double AVGMarkEGE_Quota =
                                (from ev in EV
                                 join mark in context.Mark on ev.AbiturientId equals mark.AbiturientId
                                 join ent in context.ExamInEntry on ev.EntryId equals ent.EntryId
                                 join exam in context.Exam on ent.ExamId equals exam.Id 
                                 where ev.IsCel == 0 && ev.IsQuota == 1 && ev.IsCrimea == 0 && ev.IsBE == 0
                                 && exam.IsAdditional == false
                                 select (int)mark.Value).DefaultIfEmpty(0).Average();
                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p7_2", ""));
                            node.InnerText = (AVGMarkEGE_Quota).ToString(); 

                        //р8 Проходной балл по направлению подготовки (специальности), приведенный к 100-бальной шкале
                            double MinMark =
                                (from ev in EV
                                 join mark in context.Mark on ev.AbiturientId equals mark.AbiturientId
                                 where ev.IsCel == 0 && ev.IsQuota == 0 && ev.IsCrimea == 0 && ev.IsBE == 0
                                 select new { mark.Value, mark.AbiturientId })
                                 .GroupBy(x => x.AbiturientId)
                                 .Select(x => x.Where(y => y.AbiturientId == x.Key).Select(y => (int)y.Value).DefaultIfEmpty(0).Average()
                                 ).DefaultIfEmpty(0d).Min();
                            node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p8", ""));
                            node.InnerText = (MinMark).ToString();
                    }

                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "XML files|*.xml";
                    if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        XmlWriter w = XmlWriter.Create(sfd.FileName, new XmlWriterSettings() { NewLineHandling = NewLineHandling.Entitize, NewLineChars = "" });
                        doc.Save(w);
                    }
                //}
                //catch (Exception ex)
                //{
                //    WinFormsServ.Error(ex);
                //}
                //finally
                //{
                    pf.Close();
                //}
                

                //retString = declaration.OuterXml + doc.InnerXml;
            }
        }
    }
}
