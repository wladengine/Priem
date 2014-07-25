using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using System.Xml;

using EducServLib;

namespace Priem
{
    public partial class MigratorFISEGE : Form
    {
        #region Dics

        /// <summary>
        /// Справочник № 1 - Общеобразовательные предметы
        /// </summary>
        private Dictionary<string, string> dic01_Subject = new Dictionary<string,string>();
        /// <summary>
        /// Справочник № 2 - Уровень образования
        /// </summary>
        private Dictionary<string, string> dic02_StudyLevel = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 3 - Уровень олимпиады
        /// </summary>
        private Dictionary<string, string> dic03_OlympLevel = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 4 - Статус заявления
        /// </summary>
        private Dictionary<string, string> dic04_ApplicationStatus = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 5 - Пол
        /// </summary>
        private Dictionary<string, string> dic05_Sex = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 6 - Основание для оценки
        /// </summary>
        private Dictionary<string, string> dic06_MarkDocument = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 7 - Страна
        /// </summary>
        private Dictionary<string, string> dic07_Country = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 10 - Направления подготовки
        /// </summary>
        private Dictionary<string, string> dic10_Direction = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 11 - Тип вступительных испытаний
        /// </summary>
        private Dictionary<string, string> dic11_Country = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 12 - Статус проверки заявлений
        /// </summary>
        private Dictionary<string, string> dic12_ApplicationCheckStatus = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 13 - Статус проверки документа
        /// </summary>
        private Dictionary<string, string> dic13_DocumentCheckStatus = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 14 - Форма обучения
        /// </summary>
        private Dictionary<string, string> dic14_EducationForm = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 15 - Источник финансирования
        /// </summary>
        private Dictionary<string, string> dic15_FinSource = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 17 - Сообщения об ошибках
        /// </summary>
        private Dictionary<string, string> dic17_Errors = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 18 - Тип диплома
        /// </summary>
        private Dictionary<string, string> dic18_DiplomaType = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 19 - Олимпиады
        /// </summary>
        private Dictionary<string, string> dic19_Olympics = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 22 - Тип документа, удостоверяющего личность
        /// </summary>
        private Dictionary<string, string> dic22_IdentityDocumentType = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 23 - Группа инвалидности
        /// </summary>
        private Dictionary<string, string> dic23_DisabilityType = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 30 - Вид льготы
        /// </summary>
        private Dictionary<string, string> dic30_BenefitKind = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 31 - Тип документа
        /// </summary>
        private Dictionary<string, string> dic31_DocumentType = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 33 - Тип документа для вступительного испытания ОУ
        /// </summary>
        private Dictionary<string, string> dic33_ = new Dictionary<string, string>();
        /// <summary>
        /// Справочник № 34 - Статус приемной кампании
        /// </summary>
        private Dictionary<string, string> dic34_CampaignStatus = new Dictionary<string, string>();
        
        #endregion

        public MigratorFISEGE()
        {
            InitializeComponent();
            FillComboStudyLevelGroup();
            FillComboFaculty();
        }

        private void FillComboStudyLevelGroup()
        {
            List<KeyValuePair<string, string>> lst = new List<KeyValuePair<string, string>>();
            lst.Add(new KeyValuePair<string, string>("1", "1 к."));
            lst.Add(new KeyValuePair<string, string>("1", "маг."));
            ComboServ.FillCombo(cbStudyLevelGroup, lst, false, false);
        }
        public int StudyLevelGroupId
        {
            get { return ComboServ.GetComboIdInt(cbStudyLevelGroup).Value; }
        }
        public int? FacultyId
        {
            get
            {
                if (cbFaculty.Text == ComboServ.DISPLAY_ALL_VALUE)
                    return null;
                else
                    return ComboServ.GetComboIdInt(cbFaculty);
            }
        }
        private void FillComboFaculty()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var vals = context.SP_Faculty.OrderBy(x => x.Id).
                    Select(x => new { x.Id, x.Name }).ToList().Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbFaculty, vals, false, true);
            }
        }
        
        public int? LicenseProgramId
        {
            get
            {
                if (cbLicenseProgram.Text == ComboServ.DISPLAY_ALL_VALUE)
                    return null;
                else
                    return ComboServ.GetComboIdInt(cbLicenseProgram);
            }
        }
        private void FillComboLicenseProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var vals = context.qEntry.Where(x => (FacultyId.HasValue ? x.FacultyId == FacultyId.Value : true) && x.StudyLevelGroupId == StudyLevelGroupId).
                    OrderBy(x => x.LicenseProgramCode).Select(x => new { x.LicenseProgramId, x.LicenseProgramCode, x.LicenseProgramName }).Distinct().ToList().
                    Select(x => new KeyValuePair<string, string>(x.LicenseProgramId.ToString(), x.LicenseProgramCode + " " + x.LicenseProgramName)).ToList();
                ComboServ.FillCombo(cbLicenseProgram, vals, false, true);
            }
        }
        private void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboLicenseProgram();
        }
        private void cbLicenseProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateGrid();
        }
        private void btnExportXML_Click(object sender, EventArgs e)
        {
            ExportXML();
        }

        private void ExportXML()
        {
            XmlDocument doc = new XmlDocument();
            //XmlImplementation imp = new XmlImplementation();

            string fname = "";

            SaveFileDialog sf = new SaveFileDialog();
            sf.Filter = "XML File|.xml";
            if (sf.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                fname = sf.FileName;
            else
                return;

            using (PriemEntities context = new PriemEntities())
            {
                //создаём корневой элемент
                doc.AppendChild(doc.CreateNode(XmlNodeType.Element, "Root", ""));

                //создаём элементы AuthData и PackageData
                XmlNode root = doc["Root"];

                root.AppendChild(doc.CreateNode(XmlNodeType.Element, "AuthData", ""));
                root.AppendChild(doc.CreateNode(XmlNodeType.Element, "PackageData", ""));

                //заполняем данные AuthData
                root = root["AuthData"];
                root.AppendChild(doc.CreateNode(XmlNodeType.Element, "Login", ""));
                root.AppendChild(doc.CreateNode(XmlNodeType.Element, "Pass", ""));

                root["Login"].InnerText = tbLogin.Text;
                root["Pass"].InnerText = tbPassword.Text;

                //----------------------------------------------------------------------------------------------------------
                //заполняем данные PackageData
                root = doc["Root"];
                root["PackageData"].AppendChild(doc.CreateNode(XmlNodeType.Element, "CampaignInfo", ""));
                root["PackageData"].AppendChild(doc.CreateNode(XmlNodeType.Element, "AdmissionInfo", ""));
                root["PackageData"].AppendChild(doc.CreateNode(XmlNodeType.Element, "Applications", ""));
                root["PackageData"].AppendChild(doc.CreateNode(XmlNodeType.Element, "OrdersOfAdmission", ""));

                //заполняем данные CampaignInfo
                root = root["PackageData"]["CampaignInfo"];
                root.AppendChild(doc.CreateNode(XmlNodeType.Element, "Campaigns", ""));
                //вкладываем внутрь текущую приёмную кампанию
                //в дальнейшем нужно вводить ВСЕ кампании за ВСЕ годы???
                root["Campaigns"].AppendChild(doc.CreateNode(XmlNodeType.Element, "Campaign", ""));
                //заполняем сведения о текущей приёмной кампании
                //Guid рандомно придуманный
                string CampaignId = new Guid("00000000-0000-0000-2013-000000000001").ToString();
                string CampaignName = "Приёмная кампания 2013";
                string YearStart = "2013";
                string YearEnd = "2013";

                root = root["Campaigns"].LastChild;

                root.AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                root["UID"].InnerText = CampaignId;
                root.AppendChild(doc.CreateNode(XmlNodeType.Element, "Name", ""));
                root["Name"].InnerText = CampaignName;
                root.AppendChild(doc.CreateNode(XmlNodeType.Element, "YearStart", ""));
                root["YearStart"].InnerText = YearStart;
                root.AppendChild(doc.CreateNode(XmlNodeType.Element, "YearEnd", ""));
                root["YearEnd"].InnerText = YearEnd;
                root.AppendChild(doc.CreateNode(XmlNodeType.Element, "StatusId", ""));
                root["StatusId"].InnerText = "1";

                root.AppendChild(doc.CreateNode(XmlNodeType.Element, "EducationForms", ""));
                root.AppendChild(doc.CreateNode(XmlNodeType.Element, "EducationLevels", ""));
                root.AppendChild(doc.CreateNode(XmlNodeType.Element, "CampaignDates", ""));

                //вводим формы обучения для данной кампании
                //данные об id форм обучения в справочнике №14
                
                root["EducationForms"].AppendChild(doc.CreateNode(XmlNodeType.Element, "EducationFormID", ""));
                root["EducationForms"].LastChild.InnerText = "";

                //вводим уровни образования
                //данные об id уровней образования в справочнике №2
                root["EducationLevels"].AppendChild(doc.CreateNode(XmlNodeType.Element, "EducationLevel", ""));

                root["EducationLevels"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Course", ""));
                root["EducationLevels"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "EducationLevelID", ""));

                var sfList = context.Entry.Where(x => 
                    (FacultyId.HasValue ? x.FacultyId == FacultyId : true) 
                    && (LicenseProgramId.HasValue ? x.LicenseProgramId == LicenseProgramId : true)).Select(x => x.StudyFormId);

                NewWatch wc = new NewWatch();

                var CampDates = context.CampaignDates.Join(context.StudyForm, x => x.StudyFormId, x => x.Id, (x, y) => new { Base = x, StudyForm = y })
                    .Join(context.StudyBasis, x => x.Base.StudyBasisId, y => y.Id, (x, y) => new { Base = x.Base, StudyForm = x.StudyForm, StudyBasis = y })
                    .Where(x => x.Base.StudyLevelId == StudyLevelGroupId && sfList.Contains(x.Base.StudyFormId ?? 0))
                    .Select(x => new
                    {
                        x.Base.Id,
                        x.Base.StudyLevelId,
                        StudyForm = x.StudyForm.Name,
                        StudyBasis = x.StudyBasis.Name,
                        x.Base.Stage,
                        x.Base.DateStart,
                        x.Base.DateEnd,
                        x.Base.DateOrder
                    }).Distinct();
                int wcMax = CampDates.Count();
                wc.SetMax(wcMax);
                
                wc.Show();
                wc.SetText("Кампании");
                foreach (var CampDate in CampDates)
                {
                    wc.PerformStep();
                    //вводим даты приёмной кампании
                    root["CampaignDates"].AppendChild(doc.CreateNode(XmlNodeType.Element, "CampaignDate", ""));

                    root["CampaignDates"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                    root["CampaignDates"].LastChild["UID"].InnerXml = CampDate.Id.ToString();

                    root["CampaignDates"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Course", ""));
                    root["CampaignDates"].LastChild["Course"].InnerXml = "1";
                    //данные об id уровней образования в справочнике №2
                    root["CampaignDates"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "EducationLevelID", ""));
                    root["CampaignDates"].LastChild["EducationLevelID"].InnerXml = CampDate.StudyLevelId.ToString();
                    //данные об id форм обучения в справочнике №14
                    root["CampaignDates"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "EducationFormID", ""));
                    root["CampaignDates"].LastChild["EducationFormID"].InnerXml = dic14_EducationForm[CampDate.StudyForm]; //CampDate.StudyFormId.ToString();
                    //данные об id источников финансирования в справочнике №15
                    root["CampaignDates"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "EducationSourceID", ""));
                    root["CampaignDates"].LastChild["EducationSourceID"].InnerXml = dic15_FinSource[CampDate.StudyBasis]; //CampDate.StudyBasisId.ToString();
                    //Этап приёмной кампании. Обязателен для 1 курса
                    root["CampaignDates"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Stage", ""));
                    root["CampaignDates"].LastChild["EducationSourceID"].InnerXml = CampDate.Stage.ToString();
                    root["CampaignDates"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DateStart", ""));
                    root["CampaignDates"].LastChild["EducationSourceID"].InnerXml = CampDate.DateStart.Value.ToShortDateString();
                    root["CampaignDates"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DateEnd", ""));
                    root["CampaignDates"].LastChild["EducationSourceID"].InnerXml = CampDate.DateEnd.Value.ToShortDateString();
                    //дата включения в приказ
                    root["CampaignDates"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DateOrder", ""));
                    root["CampaignDates"].LastChild["EducationSourceID"].InnerXml = CampDate.DateOrder.Value.ToShortDateString();
                }
                //--------------------------------------------------------------------------------------------------------------
                //заполняем данные AdmissionInfo
                //объём приёма (КЦ)
                root = doc["Root"]["PackageData"];
                root["AdmissionInfo"].AppendChild(doc.CreateNode(XmlNodeType.Element, "AdmissionVolume", ""));
                //Конкурсные группы
                root["AdmissionInfo"].AppendChild(doc.CreateNode(XmlNodeType.Element, "CompetitiveGroups", ""));

                //Заполняем объём приёма (КЦ)
                root = root["AdmissionInfo"]["AdmissionVolume"];

                var LPs = context.Entry.Where(x => (FacultyId.HasValue ? x.FacultyId == FacultyId.Value : true) 
                    && (LicenseProgramId.HasValue ? x.LicenseProgramId == LicenseProgramId.Value : true)).Select(x => x.LicenseProgramId);
                
                var PriemHelps = context.hlpLicenseProgramKCP;

                wcMax = PriemHelps.Count();
                wc.SetText("Заполняем объём приёма (КЦ)");
                wc.SetMax(wcMax);
                wc.ZeroCount();

                foreach (var p in PriemHelps.Where(x => LPs.Contains(x.LicenseProgramId)))
                {
                    wc.PerformStep();
                    root.AppendChild(doc.CreateNode(XmlNodeType.Element, "Item", ""));
                    //идентификатор объёма приёма по направлению подготовки в приёме
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                    root.LastChild["UID"].InnerXml = p.StudyLevelId.ToString() + p.LicenseProgramId.ToString();
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "CampaignUID", ""));
                    root.LastChild["CampaignUID"].InnerXml = CampaignId;
                    //Справочник №2 - Уровни образования
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "EducationLevelID", ""));
                    //root.LastChild["EducationLevelID"].InnerXml = dic02_StudyLevel[p.StudyLevel];
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Course", ""));
                    root.LastChild["Course"].InnerXml = "1";
                    //Справочник №10 - направления подготовки
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DirectionID", ""));
                    root.LastChild["DirectionID"].InnerXml = dic10_Direction[p.LicenseProgramName];
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberBuzhetO", ""));
                    root.LastChild["NumberBuzhetO"].InnerXml = p.KC_OB.ToString();
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberBuzhetOZ", ""));
                    root.LastChild["NumberBuzhetOZ"].InnerXml = p.KC_OZB.ToString();
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberBuzhetZ", ""));
                    root.LastChild["NumberBuzhetZ"].InnerXml = "";
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberPaidO", ""));
                    root.LastChild["NumberPaidO"].InnerXml = p.KC_OP.ToString();
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberPaidOZ", ""));
                    root.LastChild["NumberPaidOZ"].InnerXml = p.KC_OZP.ToString();
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberPaidZ", ""));
                    root.LastChild["NumberPaidZ"].InnerXml = "";
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberTargetO", ""));
                    root.LastChild["NumberTargetO"].InnerXml = p.KC_O_CEL.ToString();
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberTargetOZ", ""));
                    root.LastChild["NumberTargetOZ"].InnerXml = p.KC_OZ_CEL.ToString();
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberTargetZ", ""));
                    root.LastChild["NumberTargetZ"].InnerXml = "";
                }

                //Заполняем конкурсные группы
                root = root.ParentNode["CompetitiveGroups"];
                var CompetitionGroups = context.extCompetitiveGroup.Where(x => 
                    (LicenseProgramId.HasValue ? x.LicenseProgramId == LicenseProgramId.Value : true)
                    && (FacultyId.HasValue ? x.FacultyId == FacultyId.Value : true)
                    && (StudyLevelGroupId == 2 ? x.StudyLevelId == 17 : (x.StudyLevelId == 16 || x.StudyLevelId == 18)));

                wcMax = CompetitionGroups.Count();
                wc.SetText("Заполняем конкурсные группы");
                wc.SetMax(wcMax);
                wc.ZeroCount();

                foreach (var CompetitionGroup in CompetitionGroups)
                {
                    wc.PerformStep();
                    root.AppendChild(doc.CreateNode(XmlNodeType.Element, "CompetitiveGroup", ""));

                    //идентификатор конкурсной группы в приёме
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                    root.LastChild["UID"].InnerXml = CompetitionGroup.Id.ToString();
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "CampaignUID", ""));
                    root.LastChild["CampaignUID"].InnerXml = CampaignId;
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Course", ""));
                    root.LastChild["Course"].InnerXml = "1";
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Name", ""));
                    root.LastChild["Name"].InnerXml = CompetitionGroup.Name;

                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Items", ""));
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "TargetOrganizations", ""));
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "CommonBenefit", ""));
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "EntranceTestItems", ""));

                    //добавляем направление подготовки конкурсной группы
                    root.LastChild["Items"].AppendChild(doc.CreateNode(XmlNodeType.Element, "CompetitiveGroupItem", ""));
                    root.LastChild["Items"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "CompetitiveGroupItem", ""));

                    //заполняем данные о направлении подготовки для конкурсной группы
                    //это сделано для ВУЗов, где ведётся один конкурс на НЕСКОЛЬКО направлений
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"].AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"]["UID"].InnerXml = CompetitionGroup.Id.ToString();//возьмём тот же UID, что и для всей группы
                    //Справочник №2 - Уровни образования
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"].AppendChild(doc.CreateNode(XmlNodeType.Element, "EducationLevelID", ""));
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"]["EducationLevelID"].InnerXml = dic02_StudyLevel[CompetitionGroup.StudyLevelName];
                    //id направления подготовки из справочника №10
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"].AppendChild(doc.CreateNode(XmlNodeType.Element, "DirectionID", ""));
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"]["DirectionID"].InnerXml = dic10_Direction[CompetitionGroup.LicenseProgramName]; //CompetitionGroup.LicenseProgramCode; //пока пусть будет код направления
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"].AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberBudzhetO", ""));
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"]["NumberBudzhetO"].InnerXml = CompetitionGroup.KCP_OB.ToString();
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"].AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberBudzhetOZ", ""));
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"]["NumberBudzhetOZ"].InnerXml = CompetitionGroup.KCP_OZB.ToString();
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"].AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberBudzhetZ", ""));
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"]["NumberBudzhetZ"].InnerXml = "";
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"].AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberPaidO", ""));
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"]["NumberPaidO"].InnerXml = CompetitionGroup.KCP_OP.ToString();
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"].AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberPaidOZ", ""));
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"]["NumberPaidOZ"].InnerXml = CompetitionGroup.KCP_OZP.ToString();
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"].AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberPaidZ", ""));
                    root.LastChild["Items"].LastChild["CompetitiveGroupItem"]["NumberPaidZ"].InnerXml = "";

                    if ((CompetitionGroup.KCP_Cel_B ?? 0) > 0)
                    {
                        //добавляем сведения о целевом наборе (необяз)
                        root.LastChild["TargetOrganizations"].AppendChild(doc.CreateNode(XmlNodeType.Element, "TargetOrganization", ""));

                        root.LastChild["TargetOrganizations"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                        root.LastChild["TargetOrganizations"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "TargetOrganizationName", ""));
                        root.LastChild["TargetOrganizations"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Items", ""));


                        //направления подготовки целевого приёма
                        root.LastChild["TargetOrganizations"].LastChild["Items"].AppendChild(doc.CreateNode(XmlNodeType.Element, "CompetitiveGroupTargetItem", ""));

                        root.LastChild["TargetOrganizations"].LastChild["Items"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                        root.LastChild["TargetOrganizations"].LastChild["Items"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "EducationLevelID", ""));
                        root.LastChild["TargetOrganizations"].LastChild["Items"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberTargetO", ""));
                        root.LastChild["TargetOrganizations"].LastChild["Items"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberTargetOZ", ""));
                        root.LastChild["TargetOrganizations"].LastChild["Items"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "NumberTargetZ", ""));
                        //id направления подготовки из справочника №10
                        root.LastChild["TargetOrganizations"].LastChild["Items"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DirectionID", ""));
                    }

                    //добавляем "условия предоставления общей льготы" (б/э)
                    root.LastChild["CommonBenefit"].AppendChild(doc.CreateNode(XmlNodeType.Element, "CommonBenefitItem", ""));

                    root.LastChild["CommonBenefit"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                    //типы дипломов олимпиад
                    root.LastChild["CommonBenefit"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "OlympicDiplomTypes", ""));
                    root.LastChild["CommonBenefit"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "BenefitKindID", ""));
                    root.LastChild["CommonBenefit"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "IsForAllOlympics", ""));
                    //перечень олимпиад, для которых действует льгота
                    root.LastChild["CommonBenefit"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Olympics", ""));

                    //заполняем список типов дипломов олимпиад
                    //id типа диплома (справочник №18)
                    root.LastChild["CommonBenefit"].LastChild["OlympicDiplomTypes"].AppendChild(doc.CreateNode(XmlNodeType.Element, "OlimpicDiplomTypeID", ""));

                    //заполняем перечень олимпиад, для которых действует льгота
                    //id олимпиады (справочник №20 "Названия олимпиад")
                    root.LastChild["CommonBenefit"].LastChild["Olympics"].AppendChild(doc.CreateNode(XmlNodeType.Element, "OlympicID", ""));

                    var EntryExams = (from ExInEnt in context.extExamInEntry
                                      join qComp in context.qEntryToCompetitiveGroup
                                      on ExInEnt.EntryId equals qComp.EntryId
                                      where qComp.CompetitiveGroupId == CompetitionGroup.Id
                                      select new
                                      {
                                          UID = ExInEnt.ExamId,
                                          MinScore = ExInEnt.EgeMin,
                                          SubjectName = ExInEnt.ExamName
                                      }).Distinct();
                    foreach (var EntryExam in EntryExams)
                    {
                        //добавляем вступительные испытания для конкурсной группы
                        root.LastChild["EntranceTestItems"].AppendChild(doc.CreateNode(XmlNodeType.Element, "EntranceTestItem", ""));

                        root.LastChild["EntranceTestItems"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                        root.LastChild["EntranceTestItems"].LastChild["UID"].InnerXml = EntryExam.UID.ToString();
                        //вид вступительного испытания (справочник №11 "Тип вступительных испытаний")
                        root.LastChild["EntranceTestItems"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "EntranceTestTypeID", ""));
                        //форма вступительного испытания
                        root.LastChild["EntranceTestItems"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Form", ""));
                        root.LastChild["EntranceTestItems"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "MinScore", ""));
                        root.LastChild["EntranceTestItems"].LastChild["MinScore"].InnerXml = EntryExam.MinScore.ToString();

                        //название вступительного испытания
                        root.LastChild["EntranceTestItems"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "EntranceTestSubject", ""));
                        root.LastChild["EntranceTestItems"].LastChild["EntranceTestSubject"].AppendChild(doc.CreateNode(XmlNodeType.Element, "SubjectName", ""));
                        root.LastChild["EntranceTestItems"].LastChild["EntranceTestSubject"]["SubjectName"].InnerXml = EntryExam.SubjectName;

                        //условия предоставления льгот
                        root.LastChild["EntranceTestItems"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "EntranceTestBenefits", ""));

                        root.LastChild["EntranceTestItems"].LastChild["EntranceTestBenefits"].AppendChild(doc.CreateNode(XmlNodeType.Element, "EntranceTestBenefitItem", ""));
                        //заносим условие предоставления льгот
                        root.LastChild["EntranceTestItems"].LastChild["EntranceTestBenefits"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                        root.LastChild["EntranceTestItems"].LastChild["EntranceTestBenefits"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "OlympicDiplomTypes", ""));
                        root.LastChild["EntranceTestItems"].LastChild["EntranceTestBenefits"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "BenefitKindID", ""));
                        root.LastChild["EntranceTestItems"].LastChild["EntranceTestBenefits"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "IsForAllOlympics", ""));
                        root.LastChild["EntranceTestItems"].LastChild["EntranceTestBenefits"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Olympics", ""));
                        //заполняем дипломы (по справочнику №18 "Тип диплома")
                        root.LastChild["EntranceTestItems"].LastChild["EntranceTestBenefits"].LastChild["OlympicDiplomTypes"].AppendChild(doc.CreateNode(XmlNodeType.Element, "OlympicDiplomTypeID", ""));
                        //заполняем олимпиады (по справочнику №20 "Олимпиады")
                        root.LastChild["EntranceTestItems"].LastChild["EntranceTestBenefits"].LastChild["Olympics"].AppendChild(doc.CreateNode(XmlNodeType.Element, "OlympicID", ""));
                    }
                }
                //--------------------------------------------------------------------------------------------------------------
                //заполняем данные Applications

                var apps = (from abit in context.extAbit
                            join person in context.Person
                            on abit.PersonId equals person.Id
                            join compGroup in context.qEntryToCompetitiveGroup
                            on abit.EntryId equals compGroup.EntryId
                            where FacultyId.HasValue ? abit.FacultyId == FacultyId.Value : true
                            && LicenseProgramId.HasValue ? abit.LicenseProgramId == LicenseProgramId.Value : true
                            && abit.StudyLevelGroupId == StudyLevelGroupId
                            select new
                            {
                                AppUID = abit.Id,
                                ApplicationNumber = abit.RegNum,
                                RegistrationDate = abit.DocInsertDate,
                                LastDenyDate = abit.BackDocDate,
                                NeedHostel = person.Person_AdditionalInfo.HostelEduc,
                                EntrantUID = abit.PersonId,
                                EntrantSurname = abit.Surname,
                                EntrantName = abit.Name,
                                EntrantMiddleName = abit.SecondName,
                                AddInfo = person.Person_AdditionalInfo.ExtraInfo,
                                EgeDocOrigin = abit.HasOriginals,
                                person.PassportSeries, person.PassportNumber, person.PassportAuthor,
                                PassportDate = person.PassportDate,
                                abit.HasOriginals, person.NationalityId, person.PassportTypeId,
                                person.BirthDate, person.BirthPlace,
                                person.Person_EducationInfo.SchoolTypeId,
                                EducDocSeries = person.Person_EducationInfo.SchoolTypeId == 1 ? person.Person_EducationInfo.AttestatSeries : person.Person_EducationInfo.DiplomSeries,
                                EducDocRegion = person.Person_EducationInfo.SchoolTypeId == 1 ? person.Person_EducationInfo.AttestatRegion : "",
                                EducDocNum = person.Person_EducationInfo.SchoolTypeId == 1 ? person.Person_EducationInfo.AttestatNum : person.Person_EducationInfo.DiplomNum,
                                compGroup.CompetitiveGroupId, compGroup.CompetitiveGroupName
                            });
                
                root = doc["Root"]["PackageData"]["Applications"];

                wcMax += apps.Count();
                wc.SetText("Заполняем данные Applications");
                wc.SetMax(wcMax);
                wc.ZeroCount();

                foreach (var app in apps)
                {
                    wc.PerformStep();
                    root.AppendChild(doc.CreateNode(XmlNodeType.Element, "Application", ""));

                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                    root.LastChild["UID"].InnerText = app.AppUID.ToString();

                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "ApplicationNumber", ""));
                    root.LastChild["ApplicationNumber"].InnerText = app.ApplicationNumber.ToString();

                    //данные о человеке
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Entrant", ""));
                    //дата регистрации заявления в ИС
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "RegistrationDate", ""));
                    root.LastChild["RegistrationDate"].InnerText = app.RegistrationDate.HasValue ? app.RegistrationDate.Value.ToShortDateString() : "";

                    //дата отзыва заявления (если была)
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "LastDenyDate", ""));
                    root.LastChild["LastDenyDate"].InnerText = app.LastDenyDate.HasValue ? app.LastDenyDate.Value.ToShortDateString() : "";

                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "NeedHostel", ""));
                    root.LastChild["NeedHostel"].InnerText = app.NeedHostel.ToString();

                    //статус заявления (справочник ???)
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "StatusId", ""));
                    
                    //конкурсные группы для заявления
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "SelectedCompetitiveGroups", ""));
                    root.LastChild["SelectedCompetitiveGroups"].AppendChild(doc.CreateNode(XmlNodeType.Element, "CompetitiveGroupID", ""));
                    root.LastChild["SelectedCompetitiveGroups"]["CompetitiveGroupID"].InnerXml = app.CompetitiveGroupId.ToString();
                    
                    //элементы конкурсных групп для заявления
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "SelectedCompetitiveGroupItems", ""));
                    root.LastChild["SelectedCompetitiveGroupItems"].AppendChild(doc.CreateNode(XmlNodeType.Element, "CompetitiveGroupItemID", ""));
                    root.LastChild["SelectedCompetitiveGroupItems"]["CompetitiveGroupItemID"].InnerXml = app.CompetitiveGroupId.ToString();
                    
                    //формы обучения и источники финансирования выбранные абитуриентом
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "FinSourceAndEduForms", ""));
                    //общая льгота, предоставленная абитуриенту (необяз)
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "ApplicationCommonBenefit", ""));
                    //рез-ты вступительных испытаний
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "EntranceTestResults", ""));
                    //док-ты, приложенные к заявлению
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "ApplicationDocuments", ""));

                    //заполняем данные об абитуриенте
                    root.LastChild["Entrant"].AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                    root.LastChild["Entrant"]["UID"].InnerText = app.EntrantUID.ToString();
                    //Имя
                    root.LastChild["Entrant"].AppendChild(doc.CreateNode(XmlNodeType.Element, "FirstName", ""));
                    root.LastChild["Entrant"]["FirstName"].InnerText = app.EntrantName.ToString();
                    //Отчество
                    root.LastChild["Entrant"].AppendChild(doc.CreateNode(XmlNodeType.Element, "MiddleName", ""));
                    root.LastChild["Entrant"]["MiddleName"].InnerText = app.EntrantMiddleName.ToString();
                    //Фамилия
                    root.LastChild["Entrant"].AppendChild(doc.CreateNode(XmlNodeType.Element, "SecondName", ""));
                    root.LastChild["Entrant"]["SecondName"].InnerText = app.EntrantSurname.ToString();
                    //Пол (справочник???)
                    root.LastChild["Entrant"].AppendChild(doc.CreateNode(XmlNodeType.Element, "GenderID", ""));
                    //доп. сведения, представленные абитуриентом
                    root.LastChild["Entrant"].AppendChild(doc.CreateNode(XmlNodeType.Element, "CustomInformation", ""));
                    root.LastChild["Entrant"]["CustomInformation"].InnerText = app.AddInfo.ToString();

                    //заполняем формы обучения и источники финансирования выбранные абитуриентом (FinSourceAndEduForms)
                    root.LastChild["FinSourceAndEduForms"].AppendChild(doc.CreateNode(XmlNodeType.Element, "FinSourceEduForm", ""));

                    //id источника финансирования (справочник №15 "Источники финансирования")
                    root.LastChild["FinSourceAndEduForms"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "FinanceSourceID", ""));
                    //id формы обучения (справочник №14 "Формы обучения")
                    root.LastChild["FinSourceAndEduForms"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "EducationFormID", ""));
                    //UID организации
                    root.LastChild["FinSourceAndEduForms"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "TargetOrganizationUID", ""));

                    //добавляем сведения о льготе, предоставленной абитуриенту
                    root.LastChild["ApplicationCommonBenefit"].AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                    root.LastChild["ApplicationCommonBenefit"].AppendChild(doc.CreateNode(XmlNodeType.Element, "CompetitiveGroupID", ""));
                    //id типа документа-основания (Справочник №31 - "Тип документа") - необяз
                    root.LastChild["ApplicationCommonBenefit"].AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentTypeID", ""));
                    //документ-основание - необяз
                    root.LastChild["ApplicationCommonBenefit"].AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentReason", ""));
                    //id вида льготы (Справочник №30 - "Вид льготы")
                    root.LastChild["ApplicationCommonBenefit"].AppendChild(doc.CreateNode(XmlNodeType.Element, "BenefitKindID", ""));

                    //вводим документ-основание (если есть) - один из четырёх
                    List<int?> regards = new List<int?>() { 5, 6, 7 };//победител/призёр
                    var OlympDocs = context.extOlympiads.Where(x => x.AbiturientId == app.AppUID && regards.Contains(x.OlympValueId));
                    if (OlympDocs.Where(x => x.OlympTypeId == 2).Count() > 0)
                        //диплом победителя/призёра всероссийской олимпиады школьников
                        root.LastChild["ApplicationCommonBenefit"]["DocumentReason"].AppendChild(doc.CreateNode(XmlNodeType.Element, "OlympicTotalDocument", ""));
                    else if (OlympDocs.Where(x => x.OlympTypeId != 2).Count() > 0)
                        //диплом победителя/призёра олимпиады школьников
                        root.LastChild["ApplicationCommonBenefit"]["DocumentReason"].AppendChild(doc.CreateNode(XmlNodeType.Element, "OlympicDocument", ""));
                    
                    //основание для льготы по медицинским показаниям - ПОКА ЧТО В БАЗЕ НЕТ
                    //root.LastChild["ApplicationCommonBenefit"]["DocumentReason"].AppendChild(doc.CreateNode(XmlNodeType.Element, "MedicalDocument", ""));
                    //прочее - ПОКА ЧТО В БАЗЕ НЕТ
                    //root.LastChild["ApplicationCommonBenefit"]["DocumentReason"].AppendChild(doc.CreateNode(XmlNodeType.Element, "CustomDocument", ""));

                    //вводим данные об оценках на вступительных испытаниях
                    var abitMarks = context.qMark.Where(x => x.AbiturientId == app.AppUID);
                    foreach (var mrk in abitMarks)
                    {
                        root.LastChild["EntranceTestResults"].AppendChild(doc.CreateNode(XmlNodeType.Element, "EntranceTestResult", ""));

                        root.LastChild["EntranceTestResults"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                        root.LastChild["EntranceTestResults"].LastChild["UID"].InnerXml = mrk.Id.ToString();
                        root.LastChild["EntranceTestResults"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "ResultValue", ""));
                        root.LastChild["EntranceTestResults"].LastChild["ResultValue"].InnerXml = mrk.Value.ToString();
                        //ИД основания для оценки (документа-основания)
                        root.LastChild["EntranceTestResults"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "ResultSourceTypeID", ""));
                        root.LastChild["EntranceTestResults"].LastChild["ResultSourceTypeID"].InnerXml = mrk.ExamVedId.HasValue ? mrk.ExamVedId.Value.ToString() : "";

                        //вносим предмет (название)
                        root.LastChild["EntranceTestResults"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "EntranceTestSubject", ""));
                        root.LastChild["EntranceTestResults"].LastChild["EntranceTestSubject"].AppendChild(doc.CreateNode(XmlNodeType.Element, "SubjectName", ""));
                        root.LastChild["EntranceTestResults"].LastChild["EntranceTestSubject"]["SubjectName"].InnerXml = mrk.ExamName;

                        //ИД типа конкурсного испытания
                        root.LastChild["EntranceTestResults"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "EntranceTestTypeID", ""));
                        //ИД конкурсной группы
                        root.LastChild["EntranceTestResults"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "CompetitiveGroupID", ""));
                        root.LastChild["EntranceTestResults"].LastChild["CompetitiveGroupID"].InnerXml = app.CompetitiveGroupId.ToString();

                        ////вносим сведения об основании для оценки (OlympicDocument или OlympicTotalDocument или InstitutionDocument или EgeDocumentID) - необяз
                        //root.LastChild["EntranceTestResults"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "ResultDocument", ""));
                        //root.LastChild["EntranceTestResults"].LastChild["ResultDocument"].AppendChild(doc.CreateNode(XmlNodeType.Element, "OlympicDocument", ""));
                        //root.LastChild["EntranceTestResults"].LastChild["ResultDocument"].AppendChild(doc.CreateNode(XmlNodeType.Element, "OlympicTotalDocument", ""));
                        ////самостоятельное испытание (ведомость вступительного испытания)
                        //root.LastChild["EntranceTestResults"].LastChild["ResultDocument"].AppendChild(doc.CreateNode(XmlNodeType.Element, "InstitutionDocument", ""));
                        ////id свидетельства о рез-тах ЕГЭ, которое было приложено к заявлению
                        //root.LastChild["EntranceTestResults"].LastChild["ResultDocument"].AppendChild(doc.CreateNode(XmlNodeType.Element, "EgeDocumentID", ""));
                    }
                    //заносим документы, приложенные к заявлению
                    //Свидетельства о результатах ЕГЭ - необяз
                    root.LastChild["ApplicationDocuments"].AppendChild(doc.CreateNode(XmlNodeType.Element, "EgeDocuments", ""));
                    //док-т удостоверяющий личность
                    root.LastChild["ApplicationDocuments"].AppendChild(doc.CreateNode(XmlNodeType.Element, "IdentityDocument", ""));
                    //док-т об образовании
                    root.LastChild["ApplicationDocuments"].AppendChild(doc.CreateNode(XmlNodeType.Element, "EduDocuments", ""));
                    //военный билет - необяз
                    //root.LastChild["ApplicationDocuments"].AppendChild(doc.CreateNode(XmlNodeType.Element, "MilitaryCardDocument", ""));
                    //иные документы - необяз
                    //root.LastChild["ApplicationDocuments"].AppendChild(doc.CreateNode(XmlNodeType.Element, "CustomDocuments", ""));
                    //справка ГИА - необяз
                    //root.LastChild["ApplicationDocuments"].AppendChild(doc.CreateNode(XmlNodeType.Element, "GiaDocuments", ""));

                    var egeDocs = context.extEgeMark.Where(x => x.PersonId == app.EntrantUID).Select(x => new { x.Id, x.EgeCertificateId, x.Number, x.Year, x.EgeExamNameId, x.Value });
                    foreach (var cert in egeDocs.Select(x => new { x.EgeCertificateId, x.Year, x.Number }).Distinct())
                    {
                        //Заполняем данные о сертификатах
                        root.LastChild["ApplicationDocuments"]["EgeDocuments"].AppendChild(doc.CreateNode(XmlNodeType.Element, "EgeDocument", ""));

                        root.LastChild["ApplicationDocuments"]["EgeDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                        root.LastChild["ApplicationDocuments"]["EgeDocuments"].LastChild["UID"].InnerXml = cert.EgeCertificateId.ToString();
                        root.LastChild["ApplicationDocuments"]["EgeDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "OriginalReceived", ""));
                        root.LastChild["ApplicationDocuments"]["EgeDocuments"].LastChild["OriginalReceived"].InnerXml = app.EgeDocOrigin.ToString();
                        root.LastChild["ApplicationDocuments"]["EgeDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "OriginalReceived", ""));
                        root.LastChild["ApplicationDocuments"]["EgeDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentNumber", ""));
                        root.LastChild["ApplicationDocuments"]["EgeDocuments"].LastChild["DocumentNumber"].InnerXml = cert.Number;
                        root.LastChild["ApplicationDocuments"]["EgeDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentYear", ""));
                        root.LastChild["ApplicationDocuments"]["EgeDocuments"].LastChild["DocumentYear"].InnerXml = cert.Year;
                        root.LastChild["ApplicationDocuments"]["EgeDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Subjects", ""));

                        var exams = egeDocs.Where(x => x.EgeCertificateId == cert.EgeCertificateId).Select(x => new { x.EgeExamNameId, x.Value });
                        foreach (var exam in exams)
                        {
                            //заполняем оценки в сертификате ЕГЭ
                            root.LastChild["ApplicationDocuments"]["EgeDocuments"].LastChild["Subjects"].AppendChild(doc.CreateNode(XmlNodeType.Element, "SubjectData", ""));

                            //id дисциплины (справочник №1 - "Общеобразовательные предметы")
                            root.LastChild["ApplicationDocuments"]["EgeDocuments"].LastChild["Subjects"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "SubjectID", ""));
                            root.LastChild["ApplicationDocuments"]["EgeDocuments"].LastChild["Subjects"].LastChild["SubjectID"].InnerXml = exam.EgeExamNameId.ToString();
                            root.LastChild["ApplicationDocuments"]["EgeDocuments"].LastChild["Subjects"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Value", ""));
                            root.LastChild["ApplicationDocuments"]["EgeDocuments"].LastChild["Subjects"].LastChild["Value"].InnerXml = exam.Value.ToString();
                        }
                    }

                    //заполняем данные о док-те, удостоверяющем личность
                    //UID необяз
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"].AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"].AppendChild(doc.CreateNode(XmlNodeType.Element, "OriginalReceived", ""));
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"]["OriginalReceived"].InnerXml = app.HasOriginals.ToString();
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"].AppendChild(doc.CreateNode(XmlNodeType.Element, "OriginalReceivedDate", ""));
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"].AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentSeries", ""));
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"]["DocumentSeries"].InnerXml = app.PassportSeries;
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"].AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentNumber", ""));
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"]["DocumentNumber"].InnerXml = app.PassportNumber;
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"].AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentDate", ""));
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"]["DocumentDate"].InnerXml = app.PassportDate.ToString();
                    //кем выдан - необяз
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"].AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentOrganization", ""));
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"]["DocumentDate"].InnerXml = app.PassportAuthor;
                    //ID типа документа, удостовер личность (годится и PassportTypeId)
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"].AppendChild(doc.CreateNode(XmlNodeType.Element, "IdentityDocumentTypeID", ""));
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"]["IdentityDocumentTypeID"].InnerXml = app.PassportTypeId.ToString();
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"].AppendChild(doc.CreateNode(XmlNodeType.Element, "NationalityTypeID", ""));
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"]["NationalityTypeID"].InnerXml = app.NationalityId.ToString();
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"].AppendChild(doc.CreateNode(XmlNodeType.Element, "BirthDate", ""));
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"]["BirthDate"].InnerXml = app.BirthDate.ToString();
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"].AppendChild(doc.CreateNode(XmlNodeType.Element, "BirthPlace", ""));
                    root.LastChild["ApplicationDocuments"]["IdentityDocument"]["BirthPlace"].InnerXml = app.BirthPlace;

                    //вносим документы об образовании
                    root.LastChild["ApplicationDocuments"]["EduDocuments"].AppendChild(doc.CreateNode(XmlNodeType.Element, "EduDocument", ""));
                    //на выбор - один из
                    //все их объединяет три поля - Серия, Номер, Оригинал(да/нет)
                    XmlNode rootChild = root;
                    switch (app.SchoolTypeId)
                    {
                        case 1://Школа
                            {
                                //SchoolCertificateDocument - аттестат за 11 класс
                                root.LastChild["ApplicationDocuments"]["EduDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "SchoolCertificateDocument", ""));
                                rootChild = root.LastChild["ApplicationDocuments"]["EduDocuments"].LastChild["SchoolCertificateDocument"];
                                break;
                            }
                        case 2://ССУЗ
                            {
                                //MiddleEduDiplomaDocument - диплом СПО
                                root.LastChild["ApplicationDocuments"]["EduDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "MiddleEduDiplomaDocument", ""));
                                rootChild = root.LastChild["ApplicationDocuments"]["EduDocuments"].LastChild["MiddleEduDiplomaDocument"];
                                break;
                            }
                        case 3://НПО
                            {
                                //BasicDiplomaDocument - диплом НПО
                                root.LastChild["ApplicationDocuments"]["EduDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "BasicDiplomaDocument", ""));
                                rootChild = root.LastChild["ApplicationDocuments"]["EduDocuments"].LastChild["BasicDiplomaDocument"];
                                break;
                            }
                        case 4://ВУЗ
                            {
                                //HighEduDiplomaDocument - диплом ВПО
                                root.LastChild["ApplicationDocuments"]["EduDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "HighEduDiplomaDocument", ""));
                                rootChild = root.LastChild["ApplicationDocuments"]["EduDocuments"].LastChild["HighEduDiplomaDocument"];
                                break;
                            }
                        case 5://СПО
                            {
                                //MiddleEduDiplomaDocument - диплом СПО
                                root.LastChild["ApplicationDocuments"]["EduDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "MiddleEduDiplomaDocument", ""));
                                rootChild = root.LastChild["ApplicationDocuments"]["EduDocuments"].LastChild["MiddleEduDiplomaDocument"];
                                break;
                            }
                            ////IncomplHighEduDiplomaDocument - диплом о неполном ВПО
                            //root.LastChild["ApplicationDocuments"]["EduDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "IncomplHighEduDiplomaDocument", ""));
                            ////AcademicDiplomaDocument - академ. справка
                            //root.LastChild["ApplicationDocuments"]["EduDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "AcademicDiplomaDocument", ""));
                            ////SchoolCertificateBasicDocument - аттестат за 9 класс
                            //root.LastChild["ApplicationDocuments"]["EduDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "SchoolCertificateBasicDocument", ""));
                    }
                    rootChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentSeries", ""));
                    rootChild["DocumentSeries"].InnerXml = (app.EducDocRegion != "" ? (app.EducDocRegion + " ") : "") + app.EducDocSeries ?? "";
                    rootChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentNumber", ""));
                    rootChild["DocumentNumber"].InnerXml = app.EducDocNum ?? "";

                    ////вносим прочие документы
                    //root.LastChild["ApplicationDocuments"]["CustomDocuments"].AppendChild(doc.CreateNode(XmlNodeType.Element, "CustomDocument", ""));

                    //root.LastChild["ApplicationDocuments"]["CustomDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                    //root.LastChild["ApplicationDocuments"]["CustomDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "OriginalReceived", ""));
                    //root.LastChild["ApplicationDocuments"]["CustomDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "OriginalReceivedDate", ""));
                    //root.LastChild["ApplicationDocuments"]["CustomDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentSeries", ""));
                    //root.LastChild["ApplicationDocuments"]["CustomDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentNumber", ""));
                    //root.LastChild["ApplicationDocuments"]["CustomDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentDate", ""));
                    ////организация, выдавшая документ
                    //root.LastChild["ApplicationDocuments"]["CustomDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentOrganization", ""));
                    ////Доп. сведения
                    //root.LastChild["ApplicationDocuments"]["CustomDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "AdditionalInfo", ""));
                    ////тип документа (справочник №31 "Тип документа")
                    //root.LastChild["ApplicationDocuments"]["CustomDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentTypeNameText", ""));

                    ////вносим справки ГИА
                    //root.LastChild["ApplicationDocuments"]["GiaDocuments"].AppendChild(doc.CreateNode(XmlNodeType.Element, "GiaDocument", ""));

                    //root.LastChild["ApplicationDocuments"]["GiaDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "UID", ""));
                    //root.LastChild["ApplicationDocuments"]["GiaDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "OriginalReceived", ""));
                    //root.LastChild["ApplicationDocuments"]["GiaDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "OriginalReceivedDate", ""));
                    //root.LastChild["ApplicationDocuments"]["GiaDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentNumber", ""));
                    //root.LastChild["ApplicationDocuments"]["GiaDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentDate", ""));
                    //root.LastChild["ApplicationDocuments"]["GiaDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DocumentOrganization", ""));
                    ////предметы
                    //root.LastChild["ApplicationDocuments"]["GiaDocuments"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Subjects", ""));

                    //root.LastChild["ApplicationDocuments"]["GiaDocuments"].LastChild["Subjects"].AppendChild(doc.CreateNode(XmlNodeType.Element, "SubjectData", ""));

                    //root.LastChild["ApplicationDocuments"]["GiaDocuments"].LastChild["Subjects"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "SubjectID", ""));
                    //root.LastChild["ApplicationDocuments"]["GiaDocuments"].LastChild["Subjects"].LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Value", ""));
                }
                
                //--------------------------------------------------------------------------------------------------------------
                //заполняем данные OrdersOfAdmission

                root = doc["Root"]["PackageData"]["OrdersOfAdmission"];

                var studs = (from abit in context.qAbitAll
                             join entryView in context.extEntryView
                             on abit.Id equals entryView.AbiturientId
                             where abit.StudyLevelGroupId == StudyLevelGroupId
                             && FacultyId.HasValue ? abit.FacultyId == FacultyId.Value : true
                             && LicenseProgramId.HasValue ? abit.LicenseProgramId == LicenseProgramId.Value : true
                             select new
                             {
                                 ApplicationNumber = abit.RegNum,
                                 RegistrationDate = abit.DocInsertDate,
                                 DirectionID = abit.LicenseProgramCode,
                                 EducationFormID = abit.StudyFormId,
                                 FinanceSourceID = abit.StudyBasisId,
                                 EducationLevelID = abit.StudyLevelId,
                                 Stage = entryView.Date < new DateTime(2012, 8, 2) ? 1 : 2
                             });

                wcMax = studs.Count();
                wc.SetMax(wcMax);
                wc.ZeroCount();
                wc.SetText("3аполняем данные OrdersOfAdmission");

                foreach (var stud in studs)
                {
                    wc.PerformStep();
                    root.AppendChild(doc.CreateNode(XmlNodeType.Element, "OrderOfAdmission", ""));
                    
                    //направление подготовки (справочник №10 "Направления подготовки")
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "DirectionID", ""));
                    root.LastChild["DirectionID"].InnerXml = stud.DirectionID;
                    //форма обучения (справочник №14 "Формы обучения")
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "EducationFormID", ""));
                    root.LastChild["EducationFormID"].InnerXml = stud.EducationFormID.ToString();
                    //источник финансирования (Справочник №15 "Источники финансирования")
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "FinanceSourceID", ""));
                    root.LastChild["FinanceSourceID"].InnerXml = stud.FinanceSourceID.ToString();
                    //уровень образования (справочник №2 "Уровень образования")
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "EducationLevelID", ""));
                    root.LastChild["EducationLevelID"].InnerXml = stud.EducationLevelID.ToString();
                    //этап конкурса - обязателен для 1 курса
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Stage", ""));
                    root.LastChild["Stage"].InnerXml = stud.Stage.ToString();
                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "IsBeneficiary", ""));

                    root.LastChild.AppendChild(doc.CreateNode(XmlNodeType.Element, "Application", ""));
                    root.LastChild["Application"].AppendChild(doc.CreateNode(XmlNodeType.Element, "ApplicationNumber", ""));
                    root.LastChild["Application"]["ApplicationNumber"].InnerXml = stud.ApplicationNumber;
                    root.LastChild["Application"].AppendChild(doc.CreateNode(XmlNodeType.Element, "RegistrationDate", ""));
                    root.LastChild["Application"]["RegistrationDate"].InnerXml = stud.RegistrationDate.ToString();
                }
                //--------------------------------------------------------------------------------------------------------------
                wc.Close();
                doc.Save(fname);
                MessageBox.Show("OK");
            }
        }
        private void UpdateGrid()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var src = from c in context.extCompetitiveGroup
                          where FacultyId.HasValue ? c.FacultyId == FacultyId : true
                          && LicenseProgramId.HasValue ? c.LicenseProgramId == LicenseProgramId : true
                          select new
                          {
                              c.Name,
                              c.KCP_OB,
                              c.KCP_OP,
                              c.KCP_OZB,
                              c.KCP_OZP
                          };
                dgvCompGroup.DataSource = src;
            }
        }
        private void UpdateDictionaries()
        {
            //обновляем по очереди словари
            XmlDocument query_doc = new XmlDocument();
            XmlDocument result_doc = new XmlDocument();
            query_doc.AppendChild(query_doc.CreateNode(XmlNodeType.Element, "Root", ""));
            query_doc["Root"].AppendChild(query_doc.CreateNode(XmlNodeType.Element, "AuthData", ""));
            query_doc["Root"].AppendChild(query_doc.CreateNode(XmlNodeType.Element, "GetDictionaryContent", ""));
            
            query_doc["Root"]["AuthData"].AppendChild(query_doc.CreateNode(XmlNodeType.Element, "Login", ""));
            query_doc["Root"]["AuthData"]["Login"].InnerXml = tbLogin.Text;
            query_doc["Root"]["AuthData"].AppendChild(query_doc.CreateNode(XmlNodeType.Element, "Pass", ""));
            query_doc["Root"]["AuthData"]["Pass"].InnerXml = tbPassword.Text;

            query_doc["Root"]["GetDictionaryContent"].AppendChild(query_doc.CreateNode(XmlNodeType.Element, "DictionaryCode", ""));
            XmlNode DictionaryCodeNode = query_doc["Root"]["GetDictionaryContent"]["DictionaryCode"];

            List<int> list_codes = new List<int>() { 1, 2, 3, 4, 5, 6, 7, 10, 11, 12, 13, 14, 15, 17, 18, 19, 22, 23, 30, 31, 33, 34 };
            
            foreach (int code in list_codes)
            {
                result_doc.InnerXml = "";
                DictionaryCodeNode.InnerXml = code.ToString();
                string q_str = query_doc.InnerXml;
                WebClient client = new WebClient();
                client.Encoding = Encoding.UTF8;
                client.Headers["Content-Type"] = "text/xml";
                try
                {
                    result_doc.LoadXml(client.UploadString("http://priem.edu.ru:8000/import/importservice.svc/dictionarydetails", q_str));
                }
                catch (WebException ex)
                {
                    MessageBox.Show(ex.Message);
                }
                UpdateDictionary(code, ref result_doc);
            }
        }

        private void UpdateDictionary(int dicCode, ref XmlDocument xmlData)
        {
            if (string.IsNullOrEmpty(xmlData.InnerXml))
                return;
            
            Dictionary<string, string> dic = new Dictionary<string, string>();
            try
            {
                foreach (XmlNode node in xmlData["DictionaryItems"].ChildNodes)
                    dic.Add(node["ID"].InnerXml, node["Name"].InnerXml);
            }
            catch (Exception)
            {
                return;
            }
            switch (dicCode)
            {
                case 1: { dic01_Subject = dic; break; }
                case 2: { dic02_StudyLevel = dic; break; }
                case 3: { dic03_OlympLevel = dic; break; }
                case 4: { dic04_ApplicationStatus = dic; break; }
                case 5: { dic05_Sex = dic; break; }
                case 6: { dic06_MarkDocument = dic; break; }
                case 7: { dic07_Country = dic; break; }
                case 10: { dic10_Direction = dic; break; }
                case 11: { dic11_Country = dic; break; }
                case 12: { dic12_ApplicationCheckStatus = dic; break; }
                case 13: { dic13_DocumentCheckStatus = dic; break; }
                case 14: { dic14_EducationForm = dic; break; }
                case 15: { dic15_FinSource = dic; break; }
                case 17: { dic17_Errors = dic; break; }
                case 18: { dic18_DiplomaType = dic; break; }
                case 19: { dic19_Olympics = dic; break; }
                case 22: { dic22_IdentityDocumentType = dic; break; }
                case 23: { dic23_DisabilityType = dic; break; }
                case 30: { dic30_BenefitKind = dic; break; }
                case 31: { dic31_DocumentType = dic; break; }
                case 33: { dic33_ = dic; break; }
                case 34: { dic34_CampaignStatus = dic; break; }
            }
        }

        private void cbStudyLevelGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboLicenseProgram();
        }
    }
}
