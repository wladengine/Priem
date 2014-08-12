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
        }

        private int? StudyLevelId
        {
            get { return ComboServ.GetComboIdInt(cbStudyLevel); }
        }

        private void FillCombos()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = context.StudyLevel.Select(x => new { x.Id, x.Name }).ToList().Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();

                ComboServ.FillCombo(cbStudyLevel, src, false, true);
            }
        }

        private void btnStartImport_Click(object sender, EventArgs e)
        {
            XmlDocument doc = new XmlDocument();
            var rootNode = doc.AppendChild(doc.CreateNode(XmlNodeType.Element, "root", ""));
            XmlAttribute attr = doc.CreateAttribute("Id");
            attr.Value = "2609";
            rootNode.Attributes.Append(attr);
            using (PriemEntities context = new PriemEntities())
            {
                int rowNum = 0;
                var ListLP = context.Entry.Select(x => new { x.StudyFormId, x.StudyBasisId, x.LicenseProgramId, x.SP_LicenseProgram.GSGUCode }).Distinct().ToList();
                foreach (var LP in ListLP)
                {
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
                    node = rwNode.AppendChild(doc.CreateNode(XmlNodeType.Element, "fo", ""));
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

                    var ss = context.extEntryView.Where(x => x.LicenseProgramId == LP.LicenseProgramId && x.StudyFormId == LP.StudyFormId && x.StudyBasisId == LP.StudyBasisId);
                }

                //var declaration = doc.CreateXmlDeclaration("1.0", "utf-8", "");

                //retString = declaration.OuterXml + doc.InnerXml;
            }
        }
    }
}
