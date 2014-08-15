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
    public partial class StatFormGSGUForm1A : Form
    {
        public StatFormGSGUForm1A()
        {
            InitializeComponent();
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
                var src = context.StudyLevel.Where(x => x.LevelGroupId == 4).Select(x => new { x.Id, x.Name }).ToList().Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();

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
                var ListLP = context.Entry.Where(x => x.KCP.HasValue && x.KCP > 0 && (StudyLevelId.HasValue ? x.StudyLevel.Id == StudyLevelId : x.StudyLevel.LevelGroupId == 4))
                    .Select(x => new { x.StudyFormId, x.StudyBasisId, x.LicenseProgramId, x.SP_LicenseProgram.GSGUCode, x.SP_LicenseProgram.Name, x.SP_LicenseProgram.Code, x.KCP, x.KCPCel }).Distinct().ToList();
                ProgressForm pf = new ProgressForm();
                pf.SetProgressText("Загрузка данных...");
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


                    //p1_1 Количество мест для приёма граждан
                    node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p1_1", ""));
                    node.InnerText = LP.KCP.ToString();

                    //p1_2 в т.ч. целевики
                    node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p1_2", ""));
                    node.InnerText = LP.KCPCel.ToString();

                    var Abits = context.Abiturient
                        .Where(x => x.Entry.LicenseProgramId == LP.LicenseProgramId && x.Entry.StudyBasisId == LP.StudyBasisId && x.Entry.StudyFormId == LP.StudyFormId)
                        .Select(x => new { x.Id, x.CompetitionId });

                    //p2_1 количество заявлений
                    node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p2_1", ""));
                    node.InnerText = Abits.Count().ToString();

                    //p2_2 из них на целевые места
                    node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p2_2", ""));
                    node.InnerText = Abits.Where(x => x.CompetitionId == 6).Count().ToString();

                    var EV = context.extEntryView.Where(x => x.LicenseProgramId == LP.LicenseProgramId && x.StudyFormId == LP.StudyFormId && x.StudyBasisId == LP.StudyBasisId && x.IsCrimea == 0);

                    // p3-1 Всего зачисленных на места приема граждан
                    int cnt = EV.Count();
                    node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p3_1", ""));
                    node.InnerText = cnt.ToString();

                    // p3-2 из них на целевые места
                    cnt = EV.Where(x => x.IsCel == 1).Count();
                    node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "p3_2", ""));
                    node.InnerText = cnt.ToString();
                    
                    //р8 Проходной балл по направлению подготовки (специальности), приведенный к 100-бальной шкале
                    int MinMark =
                        (from ev in EV
                         join mark in context.Mark on ev.AbiturientId equals mark.AbiturientId
                         where ev.IsCel == 0 && ev.IsQuota == 0 && ev.IsCrimea == 0 && ev.IsBE == 0
                         select new { FiveGradeValue = (int?)mark.FiveGradeValue, mark.AbiturientId })
                         .GroupBy(x => x.AbiturientId)
                         .Select(x => x.Where(y => y.AbiturientId == x.Key).Select(y => y.FiveGradeValue ?? 0).Sum())
                         .Min();
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
