using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Data.Objects;
using System.Transactions;

using BaseFormsLib;
using EducServLib;

namespace Priem
{
    public partial class CardPerson : CardFromList
    { 
        private int? personBarc;
     
        private bool inEnableProtocol;
        private bool inEntryView;        

        // конструктор формы
        public CardPerson(string id, int? rowInd, BaseFormEx formOwner)
        {
            InitializeComponent();
            _Id = id;             
            tcCard = tabCard;
            
            this.formOwner = formOwner;
            if(rowInd.HasValue)
                this.ownerRowIndex = rowInd.Value; 

            InitControls();     
        }
        
        public CardPerson()
            : this(null, null, null)
        {
        }

        // конструктор формы открытие и создание в нашей базе        
        public CardPerson(string id)
            : this(id, null, null)
        {                        
        }

        protected override void ExtraInit()
        { 
            base.ExtraInit();                        
            _tableName = "ed.Person";
            
            Dgv = dgvApplications;
            personBarc = 0;  

            if (_Id != null && (MainClass.IsPasha()))
                btnSetStatusPasha.Visible = tbCommentFBSPasha.Visible = lblStatusPasha.Visible = true;
            else
                btnSetStatusPasha.Visible = tbCommentFBSPasha.Visible = lblStatusPasha.Visible = false;
             
            rbMale.Checked = true;

            gbAtt.Visible = true;
            gbDipl.Visible = false;
            chbEkvivEduc.Visible = false;
            
            chbHostelAbitYes.Checked = false;
            chbHostelAbitNo.Checked = false;

            lblHasAssignToHostel.Visible = false;
            lblHasExamPass.Visible = false;
            btnPrintHostel.Enabled = false;
            btnPrintExamPass.Enabled = false;
            btnGetAssignToHostel.Enabled = false;
            btnGetExamPass.Enabled = false; 
            
            tbNum.Enabled = false;
            tbFBSStatus.Enabled = false;  
            
            btnAddAbit.Enabled = false;

            if (_Id == null)
                tpEge.Parent = null;

            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    ComboServ.FillCombo(cbPassportType, HelpClass.GetComboListByTable("ed.PassportType", " ORDER BY Id "), false, false);
                    ComboServ.FillCombo(cbCountry, HelpClass.GetComboListByTable("ed.Country", "ORDER BY Distance, Name"), false, false);
                    ComboServ.FillCombo(cbNationality, HelpClass.GetComboListByTable("ed.Country", "ORDER BY Distance, Name"), false, false);
                    ComboServ.FillCombo(cbRegion, HelpClass.GetComboListByTable("ed.Region", "ORDER BY Distance, Name"), true, false);
                    ComboServ.FillCombo(cbRegionEduc, HelpClass.GetComboListByTable("ed.Region", "ORDER BY Distance, Name"), true, false);
                    ComboServ.FillCombo(cbLanguage, HelpClass.GetComboListByTable("ed.Language"), false, false);
                    ComboServ.FillCombo(cbCountryEduc, HelpClass.GetComboListByTable("ed.Country", "ORDER BY Distance, Name"), false, false);
                    ComboServ.FillCombo(cbMSStudyForm, HelpClass.GetComboListByTable("ed.StudyForm"), true, false);
                    ComboServ.FillCombo(cbHEStudyForm, HelpClass.GetComboListByTable("ed.StudyForm"), true, false);

                    cbSchoolCity.DataSource = context.ExecuteStoreQuery<string>("SELECT DISTINCT ed.Person_EducationInfo.SchoolCity AS Name FROM ed.Person_EducationInfo WHERE ed.Person_EducationInfo.SchoolCity > '' ORDER BY 1");
                    cbAttestatSeries.DataSource = context.ExecuteStoreQuery<string>("SELECT DISTINCT ed.Person_EducationInfo.AttestatSeries AS Name FROM ed.Person_EducationInfo WHERE ed.Person_EducationInfo.AttestatSeries > '' ORDER BY 1");
                    cbHEQualification.DataSource = context.ExecuteStoreQuery<string>("SELECT DISTINCT ed.Person_EducationInfo.HEQualification AS Name FROM ed.Person_EducationInfo WHERE NOT ed.Person_EducationInfo.HEQualification IS NULL /*AND ed.Person_EducationInfo.HEQualification > ''*/ ORDER BY 1");

                    cbAttestatSeries.SelectedIndex = -1;
                    cbSchoolCity.SelectedIndex = -1;
                    cbHEQualification.SelectedIndex = -1;
                }

                // магистратура!
                if (MainClass.dbType == PriemType.PriemMag)
                {
                    tpEge.Parent = null;
                    tpSecond.Parent = null;
                    btnDocs.Visible = true;

                    ComboServ.FillCombo(cbSchoolType, HelpClass.GetComboListByQuery("SELECT Cast(ed.SchoolType.Id as nvarchar(100)) AS Id, ed.SchoolType.Name FROM ed.SchoolType WHERE ed.SchoolType.Id = 4 ORDER BY 1"), false, false);
                    tbSchoolNum.Visible = false;
                    tbSchoolName.Width = 200;
                    lblSchoolNum.Visible = false;
                    gbAtt.Visible = false;
                    gbDipl.Visible = true;
                    chbIsExcellent.Text = "Диплом с отличием";
                    btnAttMarks.Visible = false;
                    gbSchool.Visible = false;                    

                    gbEduc.Location = new Point(11, 7);
                    gbFinishStudy.Location = new Point(11, 222);
                }
                else
                {
                    ComboServ.FillCombo(cbSchoolType, HelpClass.GetComboListByTable("ed.SchoolType", "ORDER BY 1"), false, false);
                    gbMainStudy.Visible = true;
                    btnDocs.Visible = false;        
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка при инициализации формы " + exc.Message);
            }
        }

        protected override bool IsForReadOnly()
        {
            return !(MainClass.RightsToEditCards() || MainClass.RightsSov_SovMain());
        }

        protected override void SetReadOnlyFieldsAfterFill()
        {
            base.SetReadOnlyFieldsAfterFill();                  
    
            if(_Id == null)
                btnDocs.Enabled = false;
        }
        
        #region handlers

        protected override void InitHandlers()
        {
            cbSchoolType.SelectedIndexChanged += new EventHandler(UpdateAfterSchool);
            cbCountry.SelectedIndexChanged += new EventHandler(UpdateAfterCountry);
            cbCountryEduc.SelectedIndexChanged += new EventHandler(UpdateAfterCountryEduc);
        }
        protected override void NullHandlers()
        {
            cbSchoolType.SelectedIndexChanged -= new EventHandler(UpdateAfterSchool);
            cbCountry.SelectedIndexChanged -= new EventHandler(UpdateAfterCountry);
            cbCountryEduc.SelectedIndexChanged -= new EventHandler(UpdateAfterCountryEduc);
        }        

        private void UpdateAfterSchool(object sender, EventArgs e)
        {
            if (SchoolTypeId == MainClass.educSchoolId)
            {
                gbAtt.Visible = true;
                gbDipl.Visible = false;
                tbSchoolName.Width = 217;
            }               
            else
            {
                if (SchoolTypeId == 4)
                    tbSchoolName.Width = 281;
                else
                    tbSchoolName.Width = 217;
                gbAtt.Visible = false;
                gbDipl.Visible = true;
            }
        }
        private void UpdateAfterCountry(object sender, EventArgs e)
        {
            if (CountryId == MainClass.countryRussiaId)
            {
                cbRegion.Enabled = true;
                cbRegion.SelectedItem = "нет";
            }
            else
            {
                cbRegion.Enabled = false;
                cbRegion.SelectedItem = "нет";
            }
        }
        private void UpdateAfterCountryEduc(object sender, EventArgs e)
        {
            if (CountryEducId == MainClass.countryRussiaId)           
                chbEkvivEduc.Visible = false;
            else
                chbEkvivEduc.Visible = true;
        }

        private void chbHostelAbitYes_CheckedChanged(object sender, EventArgs e)
        {
            chbHostelAbitNo.Checked = !chbHostelAbitYes.Checked;
            if (_Id != null)
                btnGetAssignToHostel.Enabled = chbHostelAbitYes.Checked;
        }
        private void chbHostelAbitNo_CheckedChanged(object sender, EventArgs e)
        {
            chbHostelAbitYes.Checked = !chbHostelAbitNo.Checked;
            if (_Id != null)
                btnGetAssignToHostel.Enabled = !chbHostelAbitNo.Checked;
        }

        #endregion

        private void FillPersonData(ref extPerson person)
        {
            CardTitle = Util.GetFIO(person.Surname, person.Name, person.SecondName);
            PersonName = person.Name;
            SecondName = person.SecondName;
            Surname = person.Surname;
            BirthDate = person.BirthDate;
            BirthPlace = person.BirthPlace;
            PassportTypeId = person.PassportTypeId;
            PassportSeries = person.PassportSeries;
            PassportNumber = person.PassportNumber;
            PassportAuthor = person.PassportAuthor;
            PassportDate = person.PassportDate;
            PassportCode = person.PassportCode;
            PersonalCode = person.PersonalCode;
            SNILS = person.SNILS;

            Sex = person.Sex;
            CountryId = person.CountryId;
            NationalityId = person.NationalityId;
            RegionId = person.RegionId;
            Phone = person.Phone;
            Mobiles = person.Mobiles;
            Email = person.Email;
            Code = person.Code;
            City = person.City;
            Street = person.Street;
            House = person.House;
            Korpus = person.Korpus;
            Flat = person.Flat;
            CodeReal = person.CodeReal;
            CityReal = person.CityReal;
            StreetReal = person.StreetReal;
            HouseReal = person.HouseReal;
            KorpusReal = person.KorpusReal;
            FlatReal = person.FlatReal;
            KladrCode = person.KladrCode;
            HostelAbit = person.HostelAbit;
            HostelEduc = person.HostelEduc;
            HasAssignToHostel = person.HasAssignToHostel;
            HostelFacultyId = person.HostelFacultyId;
            HasExamPass = person.HasExamPass;
            ExamPassFacultyId = person.ExamPassFacultyId;
            IsExcellent = person.IsExcellent;
            LanguageId = person.LanguageId;
            SchoolCity = person.SchoolCity;
            SchoolTypeId = person.SchoolTypeId;
            SchoolName = person.SchoolName;
            SchoolNum = person.SchoolNum;
            SchoolExitYear = person.SchoolExitYear;
            CountryEducId = person.CountryEducId;
            RegionEducId = person.RegionEducId;
            IsEqual = person.IsEqual;
            AttestatRegion = person.AttestatRegion;
            AttestatSeries = person.AttestatSeries;
            AttestatNum = person.AttestatNum;
            DiplomSeries = person.DiplomSeries;
            DiplomNum = person.DiplomNum;
            SchoolAVG = person.SchoolAVG;
            HighEducation = person.HighEducation;
            HEProfession = person.HEProfession;
            HEQualification = person.HEQualification;
            HEEntryYear = person.HEEntryYear;
            HEExitYear = person.HEExitYear;
            HEWork = person.HEWork;
            HEStudyFormId = person.HEStudyFormId;
            Stag = person.Stag;
            WorkPlace = person.WorkPlace;
            MSVuz = person.MSVuz;
            MSCourse = person.MSCourse;
            MSStudyFormId = person.MSStudyFormId;
            Privileges = person.Privileges;
            ExtraInfo = person.ExtraInfo;
            PersonInfo = person.PersonInfo;
            ScienceWork = person.ScienceWork;
            StartEnglish = person.StartEnglish;
            EnglishMark = person.EnglishMark;
            EgeInSpbgu = person.EgeInSPbgu;
            
            personBarc = person.Barcode;
        }

        //данные из нашей базы
        protected override void FillCard()
        {
            if (_Id == null)
                return;                   
                        
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    extPerson person = (from pr in context.extPerson
                                     where pr.Id == GuidId
                                     select pr).FirstOrDefault();                   

                    tbNum.Text = person.PersonNum;
                    FillPersonData(ref person);

                    FBSStatus = GetFBSStatus(GuidId);

                    FillApplications();                   
                    UpdateDataGridEge();
                    UpdateGridBenefits();

                    GetHasOriginals(context);
                    GetIsPaid();

                    inEnableProtocol = GetInEnableProtocol(context);
                    inEntryView = GetInEntryView(context);
                }
            }
            catch (DataException de)
            {
                WinFormsServ.Error("Ошибка при заполнении формы (DataException)" + de.Message);
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при заполнении формы " + ex.Message);
            }
        }
           
        public void FillApplications()
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    string queryOwn = string.Format("SELECT extAbit.Id AS Id, extAbit.FacultyAcr AS Факультет, extAbit.ProfessionCode + ' ' + extAbit.Profession AS Направление, " +
                                      "extAbit.ObrazProgram AS Образ_программа, extAbit.Specialization AS Профиль, " +
                                      "extAbit.StudyFormAcr AS Форма_обучения, extAbit.StudyBasisAcr AS Основа_обучения " +
                                      "FROM extAbit WHERE extAbit.BackDoc = 0 AND extAbit.PersonId = '{0}' ORDER BY 2, 3", _Id);


                    string queryAll = string.Format("SELECT AllAbits.Id AS Id, AllAbits.FacultyAcr AS Факультет, AllAbits.ProfessionCode + ' ' + AllAbits.Profession AS Направление, " +
                                      "AllAbits.ObrazProgram AS Образ_программа, AllAbits.Specialization AS Профиль, " +
                                      "AllAbits.StudyFormAcr AS Форма_обучения, AllAbits.StudyBasisAcr AS Основа_обучения " +
                                      "FROM AllAbits WHERE AllAbits.BackDoc = 0 AND AllAbits.PersonId = '{0}' " +
                                      "EXCEPT {1} ORDER BY 2, 3", _Id, queryOwn);

                    var sourceOwn = from abit in context.qAbiturient
                                    where !abit.BackDoc && abit.PersonId == GuidId
                                    && abit.StudyLevelGroupId == MainClass.studyLevelGroupId
                                    && abit.IsGosLine == false
                                    orderby abit.FacultyAcr, abit.ObrazProgramCrypt
                                    select new
                                    {
                                        abit.Id,
                                        Факультет = abit.FacultyAcr,
                                        Направление = abit.LicenseProgramName,
                                        Образ_программа = abit.ObrazProgramCrypt,
                                        Образ_программа_шифр = abit.ObrazProgramName,
                                        Профиль = abit.ProfileName,
                                        Форма_обучения = abit.StudyBasisName,
                                        Основа_обучения = abit.StudyFormName,
                                        abit.IsViewed
                                    };

                    var sourceAll = (from abit in context.qAbitAll
                                    where !abit.BackDoc && abit.PersonId == GuidId
                                    && abit.StudyLevelGroupId == MainClass.studyLevelGroupId
                                    orderby abit.FacultyAcr, abit.LicenseProgramName
                                    select new
                                    {
                                        abit.Id,
                                        Факультет = abit.FacultyAcr,
                                        Направление = abit.LicenseProgramName,
                                        Образ_программа = abit.ObrazProgramCrypt,
                                        Образ_программа_шифр = abit.ObrazProgramName,
                                        Профиль = abit.ProfileName,
                                        Форма_обучения = abit.StudyBasisName,
                                        Основа_обучения = abit.StudyFormName
                                    }).Except
                                    (from abit in context.qAbiturient
                                     where !abit.BackDoc && abit.PersonId == GuidId
                                     && abit.StudyLevelGroupId == MainClass.studyLevelGroupId
                                     orderby abit.FacultyAcr, abit.ObrazProgramCrypt
                                     select new
                                     {
                                         abit.Id,
                                         Факультет = abit.FacultyAcr,
                                         Направление = abit.LicenseProgramName,
                                         Образ_программа = abit.ObrazProgramCrypt,
                                         Образ_программа_шифр = abit.ObrazProgramName,
                                         Профиль = abit.ProfileName,
                                         Форма_обучения = abit.StudyBasisName,
                                         Основа_обучения = abit.StudyFormName
                                     });

                    dgvApplications.DataSource = sourceOwn;
                    dgvApplications.Columns["Id"].Visible = false;
                    dgvApplications.Columns["IsViewed"].Visible = false;
                    dgvOtherAppl.DataSource = sourceAll;
                    dgvOtherAppl.Columns["Id"].Visible = false;

                    // после зачисления раскомментить
                    var entries = (from ev in context.extEntryProtocols
                                  join ab in context.extAbit
                                  on ev.AbiturientId equals ab.Id
                                  where !ab.BackDoc && ab.PersonId == GuidId
                                  select ab).FirstOrDefault();

                    if(entries == null)                    
                        gbEnter.Visible = false;
                    else
                    {
                        gbEnter.Visible = true;
                        lblFaculty.Text = entries.FacultyName;
                        lblStudyForm.Text = entries.StudyFormName;
                        lblStudyBasis.Text = entries.StudyBasisName;
                        lblProfession.Text = entries.LicenseProgramCode + " " + entries.LicenseProgramName;
                        lblObrazProgram.Text = entries.ObrazProgramCrypt + " " + entries.ObrazProgramName;
                        lblProfile.Text = entries.ProfileId == null ? "" : entries.ProfileName;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void dgvApplications_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex >= 0 && !(bool)dgvApplications["IsViewed", e.RowIndex].Value)
            {
                e.CellStyle.BackColor = Color.LightCoral;
            }
        }
        
        // возвращает, есть ли человек в протоколе о допуске
        private bool GetInEnableProtocol(PriemEntities context)
        {  
            List<Guid> lstAbits = (from ab in context.Abiturient
                                  where ab.PersonId == GuidId
                                  select ab.Id).ToList();

            int cntProt = (from ph in context.extProtocol
                          where ph.ProtocolTypeId == 1 && !ph.IsOld && !ph.Excluded && lstAbits.Contains(ph.AbiturientId)
                          select ph.AbiturientId).Count();
            if (cntProt > 0)
                return true;
            else
                return false;     
        }
        
        // возвращает, есть ли человек в представлении к зачислению
        private bool GetInEntryView(PriemEntities context)
        {
            List<Guid> lstAbits = (from ab in context.Abiturient
                                   where ab.PersonId == GuidId
                                   select ab.Id).ToList();

            int cntProt = (from ph in context.extEntryView
                           where lstAbits.Contains(ph.AbiturientId)
                           select ph.AbiturientId).Count();
            
            if (cntProt > 0)
                return true;
            else
                return false;
        }

        BackgroundWorker bw;
        private void GetHasOriginals(PriemEntities context)
        {
            bw = new BackgroundWorker();
            bw.DoWork += bw_DoWork;
            bw.RunWorkerCompleted += bw_RunWorkerCompleted;

            var arg = new
            {
                Context = context,
                GuidId = GuidId
            };

            bw.RunWorkerAsync(arg);
            lblSearchingOriginals.Visible = true;

            //qAbiturient_WhoSetHasOriginals _who = (from orig in context.qAbiturient_WhoSetHasOriginals
            //            join Abit in context.Abiturient on orig.Id equals Abit.Id
            //            where orig.PersonId == GuidId && !Abit.BackDoc
            //            select orig).FirstOrDefault();

            //if (_who == null)
            //    return;
            //string who = _who.UserId;
            //string whoFac = _who.FacultyName;
            //string whoDate = _who.ActionTime.ToShortDateString() + " " + _who.ActionTime.ToShortTimeString();
            //who = MainClass.GetADUserName(who);

            //if (!string.IsNullOrEmpty(who))
            //{
            //    lblHasOriginalsUser.Text = "Проставлено: " + who + " (" + whoDate + " " + whoFac + ")";
            //    lblHasOriginalsUser.Visible = true;
            //    chbHasOriginals.Checked = true;
            //}
            //else
            //{
            //    lblHasOriginalsUser.Text = "";
            //    lblHasOriginalsUser.Visible = false;
            //    chbHasOriginals.Checked = false;
            //}
        }
        void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            lblSearchingOriginals.Visible = false;

            if (e.Result == null)
                return;

            var _who = ((qAbiturient_WhoSetHasOriginals)e.Result);
            string who = _who.UserId;
            string whoFac = _who.FacultyName;
            string whoDate = _who.ActionTime.ToShortDateString() + " " + _who.ActionTime.ToShortTimeString();
            who = MainClass.GetADUserName(who);

            if (!string.IsNullOrEmpty(who))
            {
                lblHasOriginalsUser.Text = "Проставлено: " + who + " (" + whoDate + " " + whoFac + ")";
                lblHasOriginalsUser.Visible = true;
                chbHasOriginals.Checked = true;
            }
            else
            {
                lblHasOriginalsUser.Text = "";
                lblHasOriginalsUser.Visible = false;
                chbHasOriginals.Checked = false;
            }
        }
        void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            using (PriemEntities context = new PriemEntities())
            {
                Guid gId = ((dynamic)(e.Argument)).GuidId;

                e.Result = (from orig in context.qAbiturient_WhoSetHasOriginals
                            join Abit in context.Abiturient on orig.Id equals Abit.Id
                            where orig.PersonId == GuidId && !Abit.BackDoc
                            select orig).FirstOrDefault();
            }
        }

        BackgroundWorker bw_ispaid;
        private void GetIsPaid()
        {
            bw_ispaid = new BackgroundWorker();
            bw_ispaid.DoWork += bw_ispaid_DoWork;
            bw_ispaid.RunWorkerCompleted += bw_ispaid_RunWorkerCompleted;

            var arg = GuidId;

            bw_ispaid.RunWorkerAsync(arg);
            lblSearchingOriginals.Visible = true;
        }
        void bw_ispaid_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            lblSearchingDogovor.Visible = false;

            if (e.Result == null)
                return;

            var _who = ((qAbiturient_WhoSetIsPaid)e.Result);
            string who = _who.UserId;
            string whoFac = _who.FacultyName;
            string whoDate = _who.ActionTime.ToShortDateString() + " " + _who.ActionTime.ToShortTimeString();
            who = MainClass.GetADUserName(who);

            if (!string.IsNullOrEmpty(who))
            {
                lblHasDogovorUser.Text = "Проставлено: " + who + " (" + whoDate + " " + whoFac + ")";
                lblHasDogovorUser.Visible = true;
                chbHasDogovor.Checked = true;
            }
            else
            {
                lblHasDogovorUser.Text = "";
                lblHasDogovorUser.Visible = false;
                chbHasDogovor.Checked = false;
            }
        }
        void bw_ispaid_DoWork(object sender, DoWorkEventArgs e)
        {
            using (PriemEntities context = new PriemEntities())
            {
                Guid gId = (Guid)(e.Argument);

                e.Result = (from paid in context.qAbiturient_WhoSetIsPaid
                            join Abit in context.Abiturient on paid.Id equals Abit.Id
                            where paid.PersonId == gId && !Abit.BackDoc
                            select paid).FirstOrDefault();
            }
        }

        #region ReadOnly & IsOpen

        // карточка открывается в режиме read-only
        protected override void  SetAllFieldsNotEnabled()
        {
            base.SetAllFieldsNotEnabled();
            
            dgvApplications.Enabled = true;
            dgvOtherAppl.Enabled = true;
            
            gbEge.Enabled = true;
            WinFormsServ.SetSubControlsEnabled(gbEge, false); 
            btnCardE.Enabled = true;           
            
            btnAttMarks.Enabled = true;           
                        
            if (HasAssignToHostel.Value && MainClass.RightsFacMain() && MainClass.HasRightsForFaculty(HostelFacultyId))
                SetBtnPrintHostelEnabled();            
            
            if (HasExamPass.Value && MainClass.RightsFacMain() && MainClass.HasRightsForFaculty(ExamPassFacultyId))
                SetBtnPrintExamPassEnabled();

            if (!IsForReadOnly() && !inEntryView)
                btnAddAbit.Enabled = true;

            btnDocs.Enabled = true;

            if (MainClass.IsPasha())            
                btnSetStatusPasha.Enabled = tbCommentFBSPasha.Enabled = true;
        }
        //убрать режим read-only
        protected override void SetAllFieldsEnabled()
        {
            base.SetAllFieldsEnabled();
            
            btnPrintHostel.Enabled = false;
            btnPrintExamPass.Enabled = false;

            if (HasAssignToHostel.Value)
            {
                chbHostelAbitYes.Enabled = chbHostelAbitNo.Enabled = false;
                btnGetAssignToHostel.Enabled = false;

                if (MainClass.RightsFacMain() && MainClass.HasRightsForFaculty(HostelFacultyId))
                    btnPrintHostel.Enabled = true;
            }
            else
            {
                if(chbHostelAbitYes.Checked)
                    btnGetAssignToHostel.Enabled = true;
                else
                    btnGetAssignToHostel.Enabled = false;
            }

            if (HasExamPass.Value)
            {
                btnGetExamPass.Enabled = false;

                if (MainClass.RightsFacMain() && MainClass.HasRightsForFaculty(ExamPassFacultyId))
                    btnPrintExamPass.Enabled = true;                    
            }
            else            
                btnGetExamPass.Enabled = true;              
          
            WinFormsServ.SetSubControlsEnabled(gbEge, true);
            //if (MainClass.IsPasha())         
            //    btnAddE.Enabled = btnRemoveE.Enabled = true;
            //else
            //    btnAddE.Enabled = btnRemoveE.Enabled = false; 
            
            btnAttMarks.Enabled = true;            

            tbNum.Enabled = false;
            tbFBSStatus.Enabled = false;
            gbEnter.Enabled = false;
        }
        // закрытие части полей в зависимости от прав
        protected override void SetReadOnlyFields()
        {
            if (MainClass.RightsFaculty())
            {
                tbName.Enabled = false;
                tbSurname.Enabled = false;
                tbSecondName.Enabled = false;

                dtBirthDate.Enabled = false;

                cbPassportType.Enabled = false;
                tbPassportAuthor.Enabled = false;
                tbPassportNumber.Enabled = false;
                tbPassportSeries.Enabled = false;
                dtPassportDate.Enabled = false;

                tbAttestatRegion.Enabled = false;
                tbAttestatNum.Enabled = false;
                cbAttestatSeries.Enabled = false;

                //tbDiplomNum.Enabled = false;
                //tbDiplomSeries.Enabled = false;

                gbPrivileges.Enabled = false;

                //временная добавка, ибо очень уж просили               
                btnAttMarks.Enabled = true;

                btnRemoveE.Enabled = false;  
            }

            if (inEnableProtocol && MainClass.RightsFaculty())
            {
                SetAllFieldsNotEnabled();

                tbMobiles.Enabled = true;
                gbStag.Enabled = true;
                gbPersonInfo.Enabled = true;

                tbDiplomNum.Enabled = true;
                tbDiplomSeries.Enabled = true;
                
                btnSaveChange.Enabled = true;
                btnClose.Enabled = true;
                btnAddAbit.Enabled = true;

                //попросили, чтобы можно было добавлять даже зачисленным в протокол о допуске
                gbEduc.Enabled = true;
                btnAttMarks.Enabled = true;                
            }

            if (inEnableProtocol && MainClass.RightsSov_SovMain_FacMain())
            {
                tbName.Enabled = false;
                tbSurname.Enabled = false;
                tbSecondName.Enabled = false;

                dtBirthDate.Enabled = false;

                cbPassportType.Enabled = false;
                tbPassportAuthor.Enabled = false;
                tbPassportNumber.Enabled = false;
                tbPassportSeries.Enabled = false;
                dtPassportDate.Enabled = false;

                tbAttestatRegion.Enabled = false;
                tbAttestatNum.Enabled = false;
                cbAttestatSeries.Enabled = false;

                //tbDiplomNum.Enabled = false;
                //tbDiplomSeries.Enabled = false;

                gbPrivileges.Enabled = false;
               
                btnRemoveE.Enabled = false; 
            }

            if (MainClass.RightsSov_SovMain())
            {
                cbPassportType.Enabled = true;
                tbPassportAuthor.Enabled = true;
                tbPassportNumber.Enabled = true;
                tbPassportSeries.Enabled = true;

                dtPassportDate.Enabled = true;

                tbAttestatRegion.Enabled = true;
                tbAttestatNum.Enabled = true;
                cbAttestatSeries.Enabled = true;

                tbDiplomNum.Enabled = true;
                tbDiplomSeries.Enabled = true;
            }

            // закрываем для создания новых для уже зачисленных
            if (inEntryView)
            {
                btnAddAbit.Enabled = false;
                chbIsExcellent.Enabled = false;
                tbSchoolAVG.Enabled = false;

                btnAddE.Enabled = false;
                btnRemoveE.Enabled = false;
            } 
        }        

        private void SetBtnPrintHostelEnabled()
        {
            gbHostel.Enabled = true;
            btnGetAssignToHostel.Enabled = false;
            btnPrintHostel.Enabled = true;
        }
        private void SetBtnPrintExamPassEnabled()
        {
            gbExamPass.Enabled = true;
            btnGetExamPass.Enabled = false;
            btnPrintExamPass.Enabled = true;
        }
        
        #endregion

        #region Save

        //Вроде как эта функция пережёвывает апострофы в понятные серверу символы. Типа как mysql_escape_string()
        private string GetOConnorString(string source)
        {
            string res = "";
            foreach (char ch in source.ToCharArray())
            {
                if (ch == '\'')
                    res += '\\' + '\'';
                else
                    res += ch;
            }
            return res;
        }

        // проверка на уникальность абитуриента
        private bool CheckIdent()
        {
            using (PriemEntities context = new PriemEntities())
            {
                ObjectParameter boolPar = new ObjectParameter("result", typeof(bool));

                if(_Id == null)
                    context.CheckPersonIdent(GetOConnorString(Surname), Name, SecondName, BirthDate, PassportSeries, PassportNumber, AttestatRegion, AttestatSeries, AttestatNum, boolPar);
                else
                    context.CheckPersonIdentWithId(GetOConnorString(Surname), Name, SecondName, BirthDate, PassportSeries, PassportNumber, AttestatRegion, AttestatSeries, AttestatNum, GuidId, boolPar);

                return Convert.ToBoolean(boolPar.Value);
            }
        }
        protected override bool CheckFields()
        { 
            // проверка на уникальность номера
            
            if (Surname.Length <= 0)
            {
                epErrorInput.SetError(tbSurname, "Отсутствует фамилия абитуриента");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epErrorInput.Clear();
            
            if (PersonName.Length <= 0)
            {
                epErrorInput.SetError(tbName, "Отсутствует имя абитуриента");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epErrorInput.Clear();

            //Для О'Коннор сделал добавку в регулярное выражение: \'
            if (!Regex.IsMatch(Surname, @"^[А-Яа-яёЁ\-\'\s]+$"))
            {
                epErrorInput.SetError(tbSurname, "Неправильный формат");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epErrorInput.Clear();

            if (!Regex.IsMatch(PersonName, @"^[А-Яа-яёЁ\-\s]+$"))
            {
                epErrorInput.SetError(tbName, "Неправильный формат");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epErrorInput.Clear();

            if (!Regex.IsMatch(SecondName, @"^[А-Яа-яёЁ\-\s]*$"))
            {
                epErrorInput.SetError(tbSecondName, "Неправильный формат");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epErrorInput.Clear();

            if (SecondName.StartsWith("-"))
            {
                SecondName = SecondName.Replace("-", "");                
            }           
            
            // проверка на англ. буквы
            if (!Util.IsRussianString(PersonName))
            {
                epErrorInput.SetError(tbName, "Имя содержит английские символы, используйте только русскую раскладку");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epErrorInput.Clear();

            if (!Util.IsRussianString(Surname))
            {
                epErrorInput.SetError(tbSurname, "Фамилия содержит английские символы, используйте только русскую раскладку");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epErrorInput.Clear();

            if (!Util.IsRussianString(SecondName))
            {
                epErrorInput.SetError(tbSecondName, "Отчество содержит английские символы, используйте только русскую раскладку");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epErrorInput.Clear();           
                       
            if (BirthDate == null)
            {
                epErrorInput.SetError(dtBirthDate, "Неправильно указана дата");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epErrorInput.Clear();

            int checkYear = DateTime.Now.Year - 12;
            if (BirthDate.Value.Year > checkYear || BirthDate.Value.Year < 1920)            
            {
                epErrorInput.SetError(dtBirthDate, "Неправильно указана дата");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epErrorInput.Clear();

            if (PassportDate.Value.Year > DateTime.Now.Year || PassportDate.Value.Year < 1970)
            {
                epErrorInput.SetError(dtPassportDate, "Неправильно указана дата");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epErrorInput.Clear();
            
            if (PassportTypeId == MainClass.pasptypeRFId)
            {
                if (!(PassportSeries.Length == 4))
                {
                    epErrorInput.SetError(tbPassportSeries, "Неправильно введена серия паспорта абитуриента");
                    tabCard.SelectedIndex = 0;
                    return false;
                }
                else
                    epErrorInput.Clear();

                if (!(PassportNumber.Length == 6))
                {
                    epErrorInput.SetError(tbPassportNumber, "Неправильно введен номер паспорта абитуриента");
                    tabCard.SelectedIndex = 0;
                    return false;
                }
                else
                    epErrorInput.Clear();
            }
            
            if (NationalityId == MainClass.countryRussiaId)
            {
                if (PassportSeries.Length <= 0)
                {
                    epErrorInput.SetError(tbPassportSeries, "Отсутствует серия паспорта абитуриента");
                    tabCard.SelectedIndex = 0;
                    return false;
                }
                else
                    epErrorInput.Clear();

                if (PassportNumber.Length <= 0)
                {
                    epErrorInput.SetError(tbPassportNumber, "Отсутствует номер паспорта абитуриента");
                    tabCard.SelectedIndex = 0;
                    return false;
                }
                else
                    epErrorInput.Clear();
            }
              
            if (PassportSeries.Length > 10)
            {
                epErrorInput.SetError(tbPassportSeries, "Слишком длинное значение серии паспорта абитуриента");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epErrorInput.Clear();
           

            if (PassportNumber.Length > 20)
            {
                epErrorInput.SetError(tbPassportNumber, "Слишком длинное значение номера паспорта абитуриента");
                tabCard.SelectedIndex = 0;
                return false;
            }
            else
                epErrorInput.Clear();

            if (!chbHostelAbitYes.Checked && !chbHostelAbitNo.Checked)
            {
                epErrorInput.SetError(chbHostelAbitNo, "Не указаны данные о предоставлении общежития");
                tabCard.SelectedIndex = 1;
                return false;
            }
            else
                epErrorInput.Clear();

            if (!chbHostelEducYes.Checked && !chbHostelEducNo.Checked)
            {
                epErrorInput.SetError(chbHostelEducNo, "Не указаны данные о предоставлении общежития");
                tabCard.SelectedIndex = 1;
                return false;
            }
            else
                epErrorInput.Clear();

            if (!Regex.IsMatch(SchoolExitYear.ToString(), @"^\d{0,4}$"))
            {
                epErrorInput.SetError(tbSchoolExitYear, "Неправильно указан год");
                tabCard.SelectedIndex = 2;
                return false;
            }
            else
                epErrorInput.Clear();

            // проверка региона аттестата - нужна ли?
            if (CountryEducId == 1)
            {
                //if (gbAtt.Visible && !Regex.IsMatch(tbAttestatRegion.Text.Trim(), @"^\w{2}$"))
                //{
                //    epErrorInput.SetError(tbAttestatRegion, "Неправильно указан регион аттестата");
                //    tabCard.SelectedIndex = 2;
                //    return false;
                //}
                //else
                //    epErrorInput.Clear();
               

                //if (gbAtt.Visible && AttestatSeries.Length <= 0)
                //{
                //    epErrorInput.SetError(cbAttestatSeries, "Отсутствует серия аттестата абитуриента");
                //    tabCard.SelectedIndex = 2;
                //    return false;
                //}
                //else
                //    epErrorInput.Clear();
            }

            if (gbAtt.Visible && AttestatNum.Length <= 0)
            {
                epErrorInput.SetError(tbAttestatNum, "Отсутствует номер аттестата абитуриента");
                tabCard.SelectedIndex = 2;
                return false;
            }
            else
                epErrorInput.Clear();

            double d = 0;
            if (tbSchoolAVG.Text.Trim() != "")
            {
                if (!double.TryParse(tbSchoolAVG.Text.Trim().Replace(".", ","), out d))
                {
                    epErrorInput.SetError(tbSchoolAVG, "Неправильный формат");
                    tabCard.SelectedIndex = 2;
                    return false;
                }
                else
                    epErrorInput.Clear();

            }

            if (MainClass.dbType == PriemType.Priem)
            {
                //if (gbDipl.Visible && tbDiplomNum.Text.Trim().Length <= 0)
                //{
                //    epErrorInput.SetError(tbDiplomNum, "Отсутствует номер диплома абитуриента");
                //    tabCard.SelectedIndex = 2;
                //    return false;
                //}
                //else
                //    epErrorInput.Clear();
            }
            
            if (!CheckIdent())
            {
                WinFormsServ.Error("В базе уже существует абитуриент с такими же либо ФИО, либо данными паспорта, либо данными аттестата!");
                return false;
            }

            return true;
        }

        protected override void InsertRec(PriemEntities context, ObjectParameter idParam)
        {
            context.Person_insert(personBarc, PersonName, SecondName, Surname, BirthDate, BirthPlace, PassportTypeId, PassportSeries, PassportNumber,
                PassportAuthor, PassportDate, Sex, CountryId, NationalityId, RegionId, Phone, Mobiles, Email,
                Code, City, Street, House, Korpus, Flat, CodeReal, CityReal, StreetReal, HouseReal, KorpusReal, FlatReal, KladrCode, HostelAbit, HostelEduc, HasAssignToHostel,
                HostelFacultyId, HasExamPass, ExamPassFacultyId, IsExcellent, LanguageId, SchoolCity, SchoolTypeId, SchoolName, SchoolNum, SchoolExitYear,
                SchoolAVG, CountryEducId, RegionEducId, IsEqual, AttestatRegion, AttestatSeries, AttestatNum, DiplomSeries, DiplomNum, HighEducation, HEProfession,
                HEQualification, HEEntryYear, HEExitYear, HEStudyFormId, HEWork, Stag, WorkPlace, MSVuz, MSCourse, MSStudyFormId, Privileges, PassportCode,
                PersonalCode, PersonInfo, ExtraInfo, ScienceWork, StartEnglish, EnglishMark, EgeInSpbgu, SNILS, idParam);
        }
        protected override void UpdateRec(PriemEntities context, Guid id)
        {
            context.Person_UpdateWithoutMain(BirthPlace, Sex, CountryId, NationalityId, RegionId, Phone, Mobiles, Email,
                Code, City, Street, House, Korpus, Flat, CodeReal, CityReal, StreetReal, HouseReal, KorpusReal, FlatReal, KladrCode, HostelAbit, HostelEduc, HasAssignToHostel,
                HostelFacultyId, HasExamPass, ExamPassFacultyId, IsExcellent, LanguageId, SchoolCity, SchoolTypeId, SchoolName, SchoolNum, SchoolExitYear,
                SchoolAVG, CountryEducId, RegionEducId, IsEqual, DiplomSeries, DiplomNum, HighEducation, HEProfession,
                HEQualification, HEEntryYear, HEExitYear, HEStudyFormId, HEWork, Stag, WorkPlace, MSVuz, MSCourse, MSStudyFormId, PassportCode,
                PersonalCode, PersonInfo, ExtraInfo, ScienceWork, StartEnglish, EnglishMark, EgeInSpbgu, id);

            if (MainClass.RightsSov_SovMain_FacMain() || MainClass.IsPasha())
                context.Person_UpdateMain(PersonName, SecondName, Surname, BirthDate, PassportTypeId, PassportSeries, PassportNumber,
                PassportAuthor, PassportDate, AttestatRegion, AttestatSeries, AttestatNum, Privileges, SNILS, id);
        }                    
                 
        protected override void OnSave()
        {
            CardTitle = Util.GetFIO(Surname, PersonName, SecondName);          
            btnAddAbit.Enabled = true;
           
            MainClass.DataRefresh();
        }

        protected override void  OnSaveNew()
        {
            using (PriemEntities context = new PriemEntities())
            {
                string num = (from pr in context.extPerson
                             where pr.Id == GuidId
                             select pr.PersonNum).FirstOrDefault().ToString();

                tbNum.Text = num;
            }
        }

        public bool IsMatchEgeNumber(string number)
        {
            string num = number.Trim();           
            if (Regex.IsMatch(num, @"^\d{2}-\d{9}-(09|10|11|12)$"))
                return true;
            else
                return false;
        }

        #endregion 

        private void tabCard_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.D1)
                this.tabCard.SelectedIndex = 0;
            if (e.Control && e.KeyCode == Keys.D2)
                this.tabCard.SelectedIndex = 1;
            if (e.Control && e.KeyCode == Keys.D3)
                this.tabCard.SelectedIndex = 2;
            if (e.Control && e.KeyCode == Keys.D4)
                this.tabCard.SelectedIndex = 3;
            if (e.Control && e.KeyCode == Keys.D5)
                this.tabCard.SelectedIndex = 4;
            if (e.Control && e.KeyCode == Keys.D6)
                this.tabCard.SelectedIndex = 5;

            if (e.Control && e.KeyCode == Keys.S)
                SaveClick();                
            if (e.Control && e.KeyCode == Keys.N)
                AddAbitClick();
        }        

        private void CardPerson_Click(object sender, EventArgs e)
        {
            this.Activate();
        }

        #region Abit

        private void dgvApplications_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
                OpenCardAbit();
        }

        private void OpenCardAbit()
        {
            if (dgvApplications.CurrentCell != null && dgvApplications.CurrentCell.RowIndex > -1)
            {
                string abId = dgvApplications.Rows[dgvApplications.CurrentCell.RowIndex].Cells["Id"].Value.ToString();
                if (abId != "")
                {
                    MainClass.OpenCardAbit(abId, this, dgvApplications.CurrentRow.Index);
                }
            }
        }

        private void AddAbitClick()
        {
            if (btnAddAbit.Visible && btnAddAbit.Enabled)
            {
                CardAbit crd = new CardAbit(GuidId);
                crd.Show();
            }
        }

        private void btnAddAbit_Click(object sender, EventArgs e)
        {
            AddAbitClick();
        }

        #endregion

        private void btnAttMarks_Click(object sender, EventArgs e)
        {
            CardAttMarks am;

            if (_Id == null)
            {              
                if (MessageBox.Show("Данное действие приведет к сохранению записи, продолжить?", "Сохранить", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    try
                    {
                        if (SaveClick())
                        {
                            am = new CardAttMarks(GuidId, !_isModified);
                            am.ShowDialog();
                        }                     
                    }
                    catch (Exception de)
                    {
                        WinFormsServ.Error("Ошибка сохранения данных" + de.Message);                        
                    }
                }
            }
            else
            {
                am = new CardAttMarks(GuidId, !_isModified);
                am.ShowDialog();
            }
        }     
               
        // Грид ЕГЭ
        #region EGE

        public void FillEgeMarks()
        {           
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    //заполнение грида с оценками
                    DataTable examTable = new DataTable();

                    DataColumn clm;
                    clm = new DataColumn();
                    clm.ColumnName = "Экзамен";
                    clm.ReadOnly = true;
                    examTable.Columns.Add(clm);

                    clm = new DataColumn();
                    clm.ColumnName = "ExamId";
                    clm.ReadOnly = true;
                    clm.DataType = typeof(int);
                    examTable.Columns.Add(clm);

                    clm = new DataColumn();
                    clm.ColumnName = "Баллы";                   
                    clm.DataType = typeof(int);
                    examTable.Columns.Add(clm);

                    clm = new DataColumn();
                    clm.ColumnName = "Зачетная";                    
                    clm.DataType = typeof(bool);
                    examTable.Columns.Add(clm);                    

                    clm = new DataColumn();
                    clm.ColumnName = "Номер_сертификата";                   
                    examTable.Columns.Add(clm);

                    IEnumerable<EgeExamName> examNames = from en in context.EgeExamName
                                                         select en;

                    foreach (EgeExamName eName in examNames)
                    {
                        DataRow newRow;
                        newRow = examTable.NewRow();
                        newRow["Экзамен"] = eName.Name;
                        newRow["ExamId"] = eName.Id;
                        examTable.Rows.Add(newRow);
                    }

                    // оценки
                    IEnumerable<extEgeMarkMax> egeMarks = from em in context.extEgeMarkMax
                                                    where em.PersonId == GuidId
                                                    select em;

                    foreach (extEgeMarkMax eMark in egeMarks)
                    {
                        for (int i = 0; i < examTable.Rows.Count; i++)
                        {
                            if (examTable.Rows[i]["ExamId"].ToString() == eMark.EgeExamNameId.ToString())
                            {                               
                                examTable.Rows[i]["Баллы"] = eMark.Value;
                                examTable.Rows[i]["Номер_сертификата"] = eMark.Number;
                                examTable.Rows[i]["Зачетная"] = eMark.IsCurrent;
                            }
                        }
                    }

                    DataView dv = new DataView(examTable);
                    dv.AllowNew = false;

                    dgvExams.DataSource = dv;
                    dgvExams.Columns["Баллы"].ValueType = typeof(int);
                    dgvExams.Columns["ExamId"].Visible = false;
                    dgvExams.ReadOnly = true;
                    dgvExams.Update();
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка  заполения грида Ege: " + exc.Message);
            }
        }

        public void UpdateDataGridEge()
        {
            if (MainClass.dbType == PriemType.PriemMag)
                return;

            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    var source = from ec in context.EgeCertificate
                                 where ec.PersonId == GuidId
                                 select new
                                 {
                                     ec.Id,
                                     ec.Number
                                 };

                    dgvEGE.DataSource = source;
                    dgvEGE.Columns["Id"].Visible = false;
                    dgvEGE.Columns["Number"].HeaderText = "Номер_сертификата";
                    dgvEGE.Columns["Number"].Width = 110;

                    btnCardE.Enabled = dgvEGE.RowCount != 0;
                    if (MainClass.RightsSov_SovMain_FacMain())
                        btnRemoveE.Enabled = dgvEGE.RowCount != 0;

                    FBSStatus = GetFBSStatus(GuidId);

                    FillEgeMarks();
                }
            }
            catch(Exception exc)
            {
                WinFormsServ.Error("Ошибка  заполения грида Ege: " + exc.Message);
            }
        }

        private void btnAddE_Click(object sender, EventArgs e)
        {
            EgeCard crd;
            if (_Id == null)
            {
                if (MessageBox.Show("Данное действие приведет к сохранению записи, продолжить?", "Сохранить", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    try
                    {
                        if (SaveClick())
                        {
                            crd = new EgeCard(GuidId);
                            crd.ToUpdateList += new UpdateListHandler(UpdateDataGridEge);
                            crd.ShowDialog();
                        }
                    }
                    catch (Exception exc)
                    {
                        WinFormsServ.Error("Ошибка сохранения данных" + exc.Message);
                    }                    
                }
            }
            else
            {
                crd = new EgeCard(GuidId);
                crd.ToUpdateList += new UpdateListHandler(UpdateDataGridEge);
                crd.ShowDialog();
            }
        }

        private void btnCardE_Click(object sender, EventArgs e)
        {
            OpenCardEge();
        }

        private void dgvEGE_DoubleClick(object sender, EventArgs e)
        {
            OpenCardEge();
        }

        private void OpenCardEge()
        {
            if (dgvEGE.CurrentCell != null && dgvEGE.CurrentCell.RowIndex > -1)
            {
                string egeId = dgvEGE.Rows[dgvEGE.CurrentCell.RowIndex].Cells["Id"].Value.ToString();
                if (egeId != "")
                {
                    EgeCard crd = new EgeCard(egeId, GuidId, GetReadOnlyEge());
                    crd.ToUpdateList += new UpdateListHandler(UpdateDataGridEge);
                    crd.ShowDialog();
                }
            }
        }

        private bool GetReadOnlyEge()
        {
//            bool isMedCollOnly = (bool)_bdc.GetValue(string.Format(@"SELECT 
//                (Case WHEN NOT EXISTS (SELECT a.Id FROM qABiturient as a 
//                INNER JOIN extEnableProtocol ON extEnableProtocol.AbiturientId = a.Id 
//                WHERE a.personid = person.id and a.FacultyId <> 11) then 1 else 0 end) AS onlyMK 
//                FROM person where Person.Id = '{0}'", _Id));

            if (!_isModified)
                return true;            
            
            if (inEnableProtocol && MainClass.RightsFaculty())
                return true;

            if (inEntryView)
                return true;

            //// закрываем уже всем на изменение кроме ограниченного набора         
            //if (!isMedCollOnly && !_bdc.RightsPashaSuper())
            //    return true;
                        
            return false;
        }

        private void btnRemoveE_Click(object sender, EventArgs e)
        {
            if (IsForReadOnly())
                return;
            
            if (MessageBox.Show("Удалить запись?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Guid certId = (Guid)dgvEGE.CurrentRow.Cells["Id"].Value;
                try
                {
                    using (PriemEntities context = new PriemEntities())
                    {
                        context.EgeCertificate_Delete(certId);
                    }                    
                }
                catch (Exception ex)
                {
                    WinFormsServ.Error("Ошибка удаления данных" + ex.Message);
                }

                UpdateDataGridEge();
            }
        }
        #endregion
        
        private void btnSetStatusPasha_Click(object sender, EventArgs e)
        {
            if (MainClass.IsPasha() || MainClass.IsOwner())
            {
                if (_Id == null)
                    return;

                if (dgvEGE.SelectedCells.Count == 0)
                {
                    WinFormsServ.Error("Нужно выбрать сертификат");
                    return;
                }

                int rwInd = dgvEGE.SelectedCells[0].RowIndex;
                Guid egeCertId = (Guid)dgvEGE["Id", rwInd].Value;

                using (PriemEntities context = new PriemEntities())
                {
                    var cert = (from ec in context.EgeCertificate
                                where ec.PersonId == GuidId && ec.Id == egeCertId && (ec.FBSStatusId == 0 || ec.FBSStatusId == 2) //&& (ec.Number.EndsWith("-00") || ec.Number.EndsWith("-10") || ec.Number.EndsWith("-11"))
                                select ec).FirstOrDefault();

                    if (cert != null)
                    {
                        if (MessageBox.Show(string.Format("Проставить статус 'Проверено' для свидетельства {0}?", cert.Number), "Внимание", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            context.EgeCertificate_UpdateFBSStatus(4, tbCommentFBSPasha.Text.Trim(), cert.Id);                           
                            MessageBox.Show("Выполнено");
                            FBSStatus = GetFBSStatus(GuidId);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Нет свидетельств, удовлетворяющих критериям");
                    }
                }
            }
        }

        private void btnDocs_Click(object sender, EventArgs e)
        {
            if (_Id == null)
                return;

            if (personBarc == 0)
            {
                MessageBox.Show("Даная анкета была заведена вручную");
                return;
            }

            if(personBarc != null)
                new DocCard(personBarc.Value, null).Show();
        }

        private void btnGetAssignToHostel_Click(object sender, EventArgs e)
        {
            using (PriemEntities context = new PriemEntities())
            {
                if (_Id == null)
                    return;

                if (HasAssignToHostel.Value)
                    return;

                int facId = MainClass.GetFacultyIdInt();

                string facName = (from fac in context.qFaculty
                                  where fac.Id == facId
                                  select fac.Name).FirstOrDefault();
               
                if (MessageBox.Show(string.Format("Будет выдано направление на поселение. Факультет: {0}, продолжить?", facName), "Сохранить", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    context.Person_UpdateHostel(true, facId, GuidId);                   
                    HasAssignToHostel = true;
                    HostelFacultyId = facId;

                    btnGetAssignToHostel.Enabled = false;

                    if (MainClass.RightsFacMain())
                        btnPrintHostel.Enabled = true;
                    
                    btnPrintHostel_Click(null, null);
                }
            }
        }

        private void btnGetExamPass_Click(object sender, EventArgs e)
        {
            using (PriemEntities context = new PriemEntities())
            {
                if (_Id == null)
                    return;

                if (HasExamPass.Value)
                    return;

                int facId = MainClass.GetFacultyIdInt();

                string facName = (from fac in context.qFaculty
                                  where fac.Id == facId
                                  select fac.Name).FirstOrDefault();                
               
                if (MessageBox.Show(string.Format("Будет выдан экзаменационный пропуск. Факультет: {0}, продолжить?", facName), "Сохранить", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {                    
                    context.Person_UpdateExamPass(true, facId, GuidId);                   
                    HasExamPass = true;
                    ExamPassFacultyId = facId;

                    btnGetExamPass.Enabled = false;

                    if (MainClass.RightsFacMain())
                        btnPrintExamPass.Enabled = true;
                    btnPrintExamPass_Click(null, null);                    
                }
            }
        }      

        private void btnPrintHostel_Click(object sender, EventArgs e)
        {
            if (_Id == null)
                return;

            sfdPrint.FileName = string.Format("{0} - направление на поселение.pdf", tbSurname.Text);
            sfdPrint.Filter = "ADOBE Pdf files|*.pdf";
            if (sfdPrint.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                Print.PrintHostelDirection(GuidId, chbPrint.Checked, sfdPrint.FileName);
        }

        private void btnPrintExamPass_Click(object sender, EventArgs e)
        {
            sfdPrint.FileName = string.Format("{0} - ЭкзПропуск.pdf", tbSurname.Text);
            sfdPrint.Filter = "ADOBE Pdf files|*.pdf";
            if (sfdPrint.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                Print.PrintExamPass(GuidId, sfdPrint.FileName, chbPrint.Checked);
        }

        private void btnAddBenefitDocument_Click(object sender, EventArgs e)
        {
            if (!GuidId.HasValue)
            {
                MessageBox.Show("Сохраните сперва карточку");
            }
            var crd = new CardBenefitDocument(null, GuidId.Value, null, this);
            crd.ToUpdateList += new UpdateListHandler(UpdateGridBenefits);
            crd.Show();
        }

        private void UpdateGridBenefits()
        {
            if (_Id == null)
                return;
            using (PriemEntities context = new PriemEntities())
            {
                var src =
                    (from PBD in context.PersonBenefitDocument
                     where PBD.PersonId == GuidId
                     select new { Id = PBD.Id, PBD.BenefitDocumentType.Name, PBD.Series, PBD.Number, PBD.HasOriginals }).ToList().
                     Select(x => new { x.Id, Name = x.Name, x.Series, x.Number, HasOriginals = x.HasOriginals ? "да" : "нет" });

                dgvBenefitDocument.DataSource = Util.ConvertToDataTable(src.ToArray());
            }
            dgvBenefitDocument.Columns["Id"].Visible = false;
            dgvBenefitDocument.Columns["Name"].HeaderText = "Тип док-та";
            dgvBenefitDocument.Columns["Series"].HeaderText = "Серия";
            dgvBenefitDocument.Columns["Number"].HeaderText = "Номер";
            dgvBenefitDocument.Columns["HasOriginals"].HeaderText = "Оригиналы";
        }

        private void btnDeleteBenefitDocument_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Удалить запись?", "Внимание", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                
            }
        }
    }
}
