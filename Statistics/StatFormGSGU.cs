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
    public partial class StatFormGSGU : Form
    {
        public StatFormGSGU()
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
                pf.MaxPrBarValue = ListLP.Count;
                pf.Show();
                try
                {
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

                        int KCP = context.Entry.Where(x => x.LicenseProgramId == LP.LicenseProgramId && x.StudyFormId == LP.StudyFormId && x.StudyBasisId == LP.StudyBasisId)
                            .Select(x => x.KCP).ToList().Select(x => x ?? 0).Sum();

                        int KCPQuota = context.Entry.Where(x => x.LicenseProgramId == LP.LicenseProgramId && x.StudyFormId == LP.StudyFormId && x.StudyBasisId == LP.StudyBasisId)
                            .Select(x => x.KCPQuota).ToList().Select(x => x ?? 0).Sum();

                        int KCPCel = context.Entry.Where(x => x.LicenseProgramId == LP.LicenseProgramId && x.StudyFormId == LP.StudyFormId && x.StudyBasisId == LP.StudyBasisId)
                            .Select(x => x.KCPCel).ToList().Select(x => x ?? 0).Sum();

                        //Всего мест для приёма граждан
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p1_1", ""));
                        node.InnerText = KCP.ToString();

                        //из них квотники
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p1_2", ""));
                        node.InnerText = KCPQuota.ToString();

                        //из них целевики
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p1_3", ""));
                        node.InnerText = KCPCel.ToString();

                        var AbitList = context.Abiturient
                            .Where(x => x.Entry.LicenseProgramId == LP.LicenseProgramId && x.Entry.StudyFormId == LP.StudyFormId && x.Entry.StudyBasisId == LP.StudyBasisId)
                            .Select(x => new { x.Id, x.CompetitionId, x.DocInsertDate });

                        //количество поданных заявлений, всего
                        int CountAbit = AbitList.Count();
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p2_1", ""));
                        node.InnerText = CountAbit.ToString();

                        //из них квотники
                        int CountAbit_VK = AbitList.Where(x => (x.CompetitionId == 2 || x.CompetitionId == 7)).Count();
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p2_2", ""));
                        node.InnerText = CountAbit_VK.ToString();

                        //из них целевики
                        int CountAbit_Cel = AbitList.Where(x => (x.CompetitionId == 6)).Count();
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p2_3", ""));
                        node.InnerText = CountAbit_Cel.ToString();

                        //из них поданные после 25.07.2014
                        int CountAbit_After2507 = AbitList.Where(x => x.DocInsertDate > new DateTime(2014, 7, 25)).Count();
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p2_4", ""));
                        node.InnerText = CountAbit_After2507.ToString();

                        var EV = context.extEntryView.Where(x => x.LicenseProgramId == LP.LicenseProgramId && x.StudyFormId == LP.StudyFormId && x.StudyBasisId == LP.StudyBasisId && x.IsCrimea == 0);

                        //зачисленных абитуриентов 31.07
                        int cnt_31072014 = EV.Where(x => x.Date < new DateTime(2014, 8, 1)).Count();
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p3_1", ""));
                        node.InnerText = cnt_31072014.ToString();

                        //зачисленных абитуриентов 05.08
                        int cnt_05082014 = EV.Where(x => x.Date < new DateTime(2014, 8, 6) && x.Date > new DateTime(2014, 8, 1)).Count();
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p3_2", ""));
                        node.InnerText = cnt_05082014.ToString();

                        //зачисленных абитуриентов 11.08
                        int cnt_11082014 = EV.Where(x => x.Date < new DateTime(2014, 8, 12) && x.Date > new DateTime(2014, 8, 7)).Count();
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p3_3", ""));
                        node.InnerText = cnt_11082014.ToString();

                        //зачисленных абитуриентов после 11.08
                        int cnt_after_11082014 = EV.Where(x => x.Date >= new DateTime(2014, 8, 12)).Count();
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p3_4", ""));
                        node.InnerText = cnt_after_11082014.ToString();

                        int cntCommonComp = EV.Where(x => x.IsCrimea == 0 && x.IsCel == 0 && x.IsQuota == 0 && x.IsBE == 0).Count();
                        int cntQuotaComp = EV.Where(x => x.IsCrimea == 0 && x.IsCel == 0 && x.IsQuota == 1 && x.IsBE == 0).Count();
                        int cntCelComp = EV.Where(x => x.IsCrimea == 0 && x.IsCel == 1 && x.IsQuota == 0 && x.IsBE == 0).Count();
                        int cntBEComp = EV.Where(x => x.IsCrimea == 0 && x.IsCel == 0 && x.IsQuota == 0 && x.IsBE == 1).Count();
                        int cntOlympLike100Balls =
                            (from exEV in context.extEntryView
                             join mrk in context.Mark on exEV.AbiturientId equals mrk.AbiturientId
                             where mrk.IsFromOlymp
                             && exEV.LicenseProgramId == LP.LicenseProgramId
                             && exEV.StudyFormId == LP.StudyFormId
                             && exEV.StudyBasisId == LP.StudyBasisId
                             && exEV.IsCrimea == 0 && exEV.IsCel == 0 && exEV.IsQuota == 0 && exEV.IsBE == 0
                             select exEV.AbiturientId).Count();

                        //зачисленных абитуриентов по общему конкурсу (без учёта получивших 100 баллов за олимпиаду)
                        int cnt_common_no_100_balls = cntCommonComp - cntOlympLike100Balls;
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p4_2", ""));
                        node.InnerText = cnt_common_no_100_balls.ToString();

                        //зачисленных абитуриентов по общему конкурсу, получивших 100 баллов за олимпиаду
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p4_3", ""));
                        node.InnerText = cntOlympLike100Balls.ToString();

                        //зачисленных абитуриентов по квоте в/к
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p4_4", ""));
                        node.InnerText = cntQuotaComp.ToString();

                        //зачисленных абитуриентов целевиков цел
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p4_5", ""));
                        node.InnerText = cntCelComp.ToString();

                        //зачисленных абитуриентов без экзаменов б/э (чемпионы Олимпийских игр (всегда = 0 для СПбГУ))
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p4_6", ""));
                        node.InnerText = "0";

                        //зачисленных абитуриентов без экзаменов б/э (всеросс. и междунар. олимпиады)
                        int cntOlympVseross =
                            (from exEV in context.extEntryView
                             join mrk in context.Olympiads on exEV.AbiturientId equals mrk.AbiturientId
                             where mrk.OlympTypeId <= 2 //всеросс и международные
                             && exEV.LicenseProgramId == LP.LicenseProgramId
                             && exEV.StudyFormId == LP.StudyFormId
                             && exEV.StudyBasisId == LP.StudyBasisId
                             && exEV.IsCrimea == 0 && exEV.IsCel == 0 && exEV.IsQuota == 0 && exEV.IsBE == 0
                             select exEV.AbiturientId).Count();
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p4_7", ""));
                        node.InnerText = cntOlympVseross.ToString();

                        //зачисленных абитуриентов без экзаменов б/э (прочие олимпиады)
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p4_8", ""));
                        node.InnerText = (cntBEComp - cntOlympVseross).ToString();

                        //зачисленных абитуриентов без экзаменов б/э (прочие олимпиады)
                        node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p5", ""));
                        node.InnerText = "";
                    }

                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "XML files|*.xml";
                    if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        XmlWriter w = XmlWriter.Create(sfd.FileName, new XmlWriterSettings() { NewLineHandling = NewLineHandling.Entitize, NewLineChars = "" });
                        doc.Save(w);
                    }
                }
                catch (Exception ex)
                {
                    WinFormsServ.Error(ex);
                }
                finally
                {
                    pf.Close();
                }
                

                //retString = declaration.OuterXml + doc.InnerXml;
            }
        }
    }
}
