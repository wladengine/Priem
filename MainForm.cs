using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Diagnostics;

using EducServLib;
using System.Threading;
using PriemLib;

namespace Priem
{
    public partial class MainForm : Form
    {
        private DBPriem _bdc;
        private string _titleString;
        private bool bSuccessAuth;

        public MainForm()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized;

            SetDB();

            try
            {
                if (string.IsNullOrEmpty(MainClass.connString))
                    return;

                bSuccessAuth = MainClass.Init(this);

                if (!bSuccessAuth)
                {
                    WinFormsServ.Error("Не удалось подключиться под вашей учетной записью");
                    return;
                }

                _bdc = MainClass.Bdc;
                string sPath = string.Format("{0}; Пользователь: {1}", _titleString, MainClass.GetUserName());
                
                OpenHelp(sPath);

                //Технические запросы к базе делаются асинхронно для ускорения запуска стартового окна
                Thread t1 = new Thread(MainClass.DeleteAllOpenByHolder);
                t1.Start();
                Thread t2 = new Thread(MainClass.InitQueryBuilder);
                t2.Start();
                Thread t3 = new Thread(ShowProtocolWarning);
                t3.Start();

                OpenStartForm();
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Не удалось подключиться под вашей учетной записью  " + exc.Message);
                msMainMenu.Enabled = false;
            }
        }

        /// <summary>
        /// Устанавливает тип приложения и строку подключения к базе
        /// </summary>
        private void SetDB()
        {
            string dbName = ConfigurationManager.AppSettings["Priem"];
            MainClass.connString = DBConstants.CS_PRIEM;
            MainClass.connStringOnline = DBConstants.CS_PriemONLINE;

            switch (dbName.ToLowerInvariant())
            {
                case "priem":
                    _titleString = " на первый курс";
                    MainClass.dbType = PriemType.Priem;
                    MainClass.IsTestDB = false;
                    break;
                case "priemsecond":
                    _titleString = " на первый курс (новые правила)";
                    MainClass.dbType = PriemType.Priem;
                    MainClass.connString = DBConstants.CS_PRIEM_SECOND;
                    MainClass.IsTestDB = false;
                    break;
                case "priemmag":
                    _titleString = " в магистратуру";
                    MainClass.dbType = PriemType.PriemMag;
                    MainClass.IsTestDB = false;
                    break;

                case "priem_fac":
                    _titleString = " рабочая 1 курс superman";
                    MainClass.connString = DBConstants.CS_PRIEM_FAC;
                    MainClass.dbType = PriemType.Priem;
                    MainClass.IsTestDB = false;
                    break;
                case "priemmag_fac":
                    _titleString = " рабочая магистратура superman";
                    MainClass.connString = DBConstants.CS_PRIEM_FAC;
                    MainClass.dbType = PriemType.PriemMag;
                    MainClass.IsTestDB = false;
                    break;
                case "priem_test":
                    _titleString = " ТЕСТОВАЯ 1 курс";
                    MainClass.connString = DBConstants.CS_PRIEM_FAC;
                    MainClass.dbType = PriemType.Priem;
                    MainClass.IsTestDB = true;
                    break;
                case "priemmag_test":
                    _titleString = " ТЕСТОВАЯ магистратура";
                    MainClass.connString = DBConstants.CS_PRIEM_FAC;
                    MainClass.dbType = PriemType.PriemMag;
                    MainClass.IsTestDB = true;
                    break;
                default:
                    WinFormsServ.Error("Проверьте параметры конфиг-файла!");
                    this.Text = "ОШИБКА";
                    return;
            }

            this.Text = "ПРИЕМ " + MainClass.sPriemYear + _titleString;
        }
        /// <summary>
        /// extra information for open - what smi are enabled
        /// </summary>
        /// <param name="path"></param>
        public void OpenHelp(string path)
        {
            try
            {
                MainClass.dirTemplates = string.Format(@"{0}\Templates", Application.StartupPath);
                tsslMain.Text = string.Format("Открыта база: Прием в СПбГУ {0} {1}; ", MainClass.sPriemYear, path);

                Thread t = new Thread(ShowMessageIfTestDB);
                t.Start();

                if (MainClass.IsOwner())
                    return;

                // магистратура!
                if (MainClass.dbType == PriemType.PriemMag)
                {
                    smiOlympAbitList.Visible = false;
                    smiOlymps.Visible = false;
                    smiOlymp2Competition.Visible = false;
                    smiOlymp2Mark.Visible = false;
                }
                else
                {
                    smiOnlineChanges.Visible = false;
                    smiLoad.Visible = false;
                }                
                
                smiRatingList.Visible = false;
                smiOrderNumbers.Visible = false;
                smiOlymps.Visible = false;
                smiCreateVed.Visible = false;
                smiBooks.Visible = false;
                smiCrypto.Visible = false;                
                smiFBS.Visible = false;
                smiExport.Visible = false;
                smiImport.Visible = false;
                smiExamsVedRoomList.Visible = false;
                smiEntryView.Visible = false;
                smiDisEntryView.Visible = false;

                smiEGEStatistics.Visible = false;
                smiDynamics.Visible = false;
                smiFormA.Visible = false;
                smiForm2.Visible = false;

                smiAbitFacultyIntesection.Visible = false;
                smiRegionStat.Visible = false;
                smiOlympStatistics.Visible = false;

                // Разделение видимости меню
                if (MainClass.IsFacMain())
                {
                    smiOlymps.Visible = true;
                    smiCreateVed.Visible = true;
                    smiExamsVedRoomList.Visible = true;
                    smiRatingList.Visible = true;
                    smiEntryView.Visible = true;
                    smiDisEntryView.Visible = true;
                    smiAbitFacultyIntesection.Visible = true;
                    smiExport.Visible = true;
                }

                if (MainClass.IsFaculty())
                {
                    smiRatingList.Visible = true;
                }

                if (MainClass.IsEntryChanger())
                {
                    smiBooks.Visible = true;
                    smiEnterManual.Visible = false;
                    smiRatingListPasha.Visible = false;
                    smiRatingList.Visible = true;
                    smiExport.Visible = true;
                }

                if (MainClass.IsPasha())
                {
                    smiCrypto.Visible = true;
                    smiBooks.Visible = true;
                    smiRatingList.Visible = true;
                    smiFBS.Visible = true;
                    smiOrderNumbers.Visible = true;
                    smiExport.Visible = true;
                    smiEntryView.Visible = true;
                    smiDisEntryView.Visible = true;
                    smiEnterManual.Visible = true;
                    smiAppeal.Visible = true;
                    smiDecryptor.Visible = true;                    

                    //Паша попросил добавить себе
                    smiCreateVed.Visible = true;
                    smiExamsVedRoomList.Visible = true;

                    smiRatingListPasha.Visible = true;

                    smiEGEStatistics.Visible = true;
                    smiDynamics.Visible = true;
                    smiFormA.Visible = true;
                    smiForm2.Visible = true;

                    smiAbitFacultyIntesection.Visible = true;
                    smiRegionStat.Visible = true;
                    smiOlympStatistics.Visible = true;
                }

                if (MainClass.IsRectorat())
                {
                    smiEGEStatistics.Visible = true;
                    smiFormA.Visible = true;

                    smiExport.Visible = true;
                    smiRatingList.Visible = true;
                    smiAbitFacultyIntesection.Visible = true;
                    smiRegionStat.Visible = true;
                    smiOlympStatistics.Visible = true;
                    smiForm2.Visible = true;
                    smiStatGSGU.Visible = true;
                }

                if (MainClass.IsSovetnik() || MainClass.IsSovetnikMain())
                {
                    smiAbitFacultyIntesection.Visible = true;
                }

                if (MainClass.IsCrypto())
                {
                    smiCrypto.Visible = true;
                    smiExamsVedRoomList.Visible = true;
                    smiAppeal.Visible = false;
                    smiDecryptor.Visible = false;
                    smiLoadMarks.Visible = false;
                }

                if (MainClass.IsCryptoMain())
                {
                    smiCrypto.Visible = true;
                    smiAppeal.Visible = true;
                    smiExamsVedRoomList.Visible = true;

                    //глава шифровалки тоже хочет создавать ведомости
                    smiCreateVed.Visible = true;
                   
                    smiDecryptor.Visible = false;
                    smiLoadMarks.Visible = false;
                }

                if (MainClass.IsPrintOrder())
                    smiEntryView.Visible = true;

                //временно                
                smiImport.Visible = false;
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc);
            }
        }
        /// <summary>
        /// Запускает стартовый лист
        /// </summary>
        private void OpenStartForm()
        {
            Form frm;
            if (MainClass._config.ValuesList.Keys.Contains("lstAbitDef"))
            {
                bool lstAbitDef = bool.Parse(MainClass._config.ValuesList["lstAbitDef"]);

                if (lstAbitDef)
                {
                    frm = new ListAbit(this);
                    smiListAbit.Checked = true;
                    smiListPerson.Checked = false;
                }
                else
                {
                    if (MainClass.dbType == PriemType.PriemMag)
                        frm = new ApplicationInetList();
                    else
                        frm = new PersonInetList();

                    smiListPerson.Checked = true;
                    smiListAbit.Checked = false;
                }
            }
            else
                frm = new PersonInetList();

            frm.Show();
        }
        /// <summary>
        /// Выводит предупреждение в случае отсутствия свежего протокола о допуске
        /// </summary>
        private void ShowProtocolWarning()
        {
            if (MainClass.dbType == PriemType.Priem && !MainClass.b1kCheckProtocolsEnabled)
                return;
            if (MainClass.dbType == PriemType.PriemMag && !MainClass.bMagCheckProtocolsEnabled)
                return;

            DateTime dtNow = DateTime.Now.Date;
            DateTime dtYesterday = DateTime.Now.AddDays(-1).Date;
            
            using (PriemEntities context = new PriemEntities())
            {
                int cntProts = (from prot in context.qProtocol
                                where prot.Date >= dtYesterday
                                select prot).Count();

                if (cntProts == 0)
                    MessageBox.Show("Уважаемые пользователи!\nВами не создан протокол о допуске!\nСоздайте срочно!\nУправление по организации приема.", "Внимание");
            }
        }

        private void ShowMessageIfTestDB()
        {
            //предупреждение об тестовом режиме базы
            if (MainClass.IsTestDB)
                MessageBox.Show("Внимание! Вы пользуетесь тестовой версией приложения, предназначенной для обучения и тестирования. Данные не сохраняются в рабочую базу", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

            ////предупреждение об рабочем режиме базы
            //MessageBox.Show("Уважаемые пользователи!\nСистема находится в рабочем режиме.\nВведение тестовых записей не допускается.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            //сохраняем параметры
            try
            {
                if (MainClass._config != null)
                {
                    MainClass._config.AddValue("lstAbitDef", smiListAbit.Checked.ToString());
                    MainClass._config.SaveConfig();
                }
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при чтении параметров из файла: " + ex.Message);
            }

            if (!bSuccessAuth)
                return;

            try
            {
                MainClass.DeleteTempFiles();
                MainClass.DeleteAllOpenByHolder();
                MainClass.SaveParameters();
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка записи в базу: " + ex.Message);
            }
        }

        //реакция на смену mdi-окна
        private void MainForm_MdiChildActivate(object sender, EventArgs e)
        {
            Form f = this.ActiveMdiChild;

            if (f is FormFilter)
                smiFilters.Visible = true;
            else
                smiFilters.Visible = false;
        }

        private void smiEntry_Click(object sender, EventArgs e)
        {
            new EntryList().Show();
        }
        private void smiLoad_Click(object sender, EventArgs e)
        {
            if (MainClass.dbType == PriemType.PriemMag)
                new ApplicationInetList().Show();
            else
                new PersonInetList().Show();
        }
        private void smiAbits_Click(object sender, EventArgs e)
        {
            new ListAbit(this).Show();
        }

        private void smiPersons_Click(object sender, EventArgs e)
        {
            new ListPersonFilter(this).Show();
        }
        private void smiAllAbitList_Click(object sender, EventArgs e)
        {
            new AllAbitList().Show();
        }
        private void smiListHostel_Click(object sender, EventArgs e)
        {
            new ListHostel().Show();
        }
        private void smiVedExamList_Click(object sender, EventArgs e)
        {
            new VedExamLists().Show();
        }

        private void smiOlymBook_Click(object sender, EventArgs e)
        {
            
        }

        private void smiEnableProtocol_Click(object sender, EventArgs e)
        {
            new EnableProtocolList().Show();
        }
        private void smiDisEnableProtocol_Click(object sender, EventArgs e)
        {
            new DisEnableProtocolList().Show();
        }
        private void smiPersonsSPBGU_Click(object sender, EventArgs e)
        {
            new PersonInetList().Show();
        }
        private void smiOnlineChanges_Click(object sender, EventArgs e)
        {
            new OnlineChangesList().Show();
        }
        private void smiOlympAbitList_Click(object sender, EventArgs e)
        {
            new OlympAbitList().Show();
        }
        private void smiExamName_Click(object sender, EventArgs e)
        {
            new ExamNameList().Show();
        }
        private void smiEGE_Click(object sender, EventArgs e)
        {
            new EgeExamList().Show();
        }
        private void smiExam_Click(object sender, EventArgs e)
        {
            new ExamList().Show();
        }
        private void smiChanges_Click(object sender, EventArgs e)
        {
            new PersonChangesList().Show();
        }
        private void smiCPK1_Click(object sender, EventArgs e)
        {
            new CPK1().Show();
        }

        #region Настройки

        //настройки по умолачнию - открывается список абитуриентов
        private void smiListPerson_Click(object sender, EventArgs e)
        {
            smiListAbit.Checked = false;
        }

        //настройки по умолачнию - открывается список заявлений
        private void smiListAbit_Click(object sender, EventArgs e)
        {
            smiListPerson.Checked = false;
        }

        //сохранить фильтр
        private void smiFiltersSave_Click(object sender, EventArgs e)
        {
            FilterList f = new FilterList(this.ActiveMdiChild as FormFilter, true);
            f.ShowDialog();
        }

        //выбрать фильтр
        private void smiFiltersChoose_Click(object sender, EventArgs e)
        {
            FilterList f = new FilterList(this.ActiveMdiChild as FormFilter, false);
            f.ShowDialog();
        }

        #endregion

        //private void импортОлимпиадToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    SomeMethodsClass.FillOlymps();
        //}

        private void smiChangeCompCel_Click(object sender, EventArgs e)
        {
            new ChangeCompCelProtocolList().Show();
        }
        private void smiExams_Click(object sender, EventArgs e)
        {
            new ExamResults().Show();
        }
        private void smiOlymp2Mark_Click(object sender, EventArgs e)
        {
            new Olymp2Mark().Show();
        }
        private void smiOlymp2Competition_Click(object sender, EventArgs e)
        {
            new Olymp2Competition().Show();
        }

        private void smiCreateVed_Click(object sender, EventArgs e)
        {
            new ExamsVedList().Show();
        }
        private void smiExamsVedRoomList_Click(object sender, EventArgs e)
        {
            new ExamsVedRoomList().Show();
        }

        private void smiMinEge_Click(object sender, EventArgs e)
        {
            new MinEgeList().Show();
        }
        private void smiHelp_Click(object sender, EventArgs e)
        {
            Process.Start(string.Format(@"{0}\Templates\Help.doc", Application.StartupPath));
        }

        private void smiChangeCompBE_Click(object sender, EventArgs e)
        {
            new ChangeCompBEProtocolList().Show();
        }
        private void smiFormA_Click(object sender, EventArgs e)
        {
            new FormA().Show();
        }
        private void smiDynamics_Click(object sender, EventArgs e)
        {
            new CountAbitStatistics().Show();
        }

        private void smiGetByFIOPasp_Click(object sender, EventArgs e)
        {
            FBSClass.MakeFBS(2);
        }
        private void smiLoadFBS_Click(object sender, EventArgs e)
        {
            if (MainClass.IsPasha())
                new LoadFBS().Show();
        }
        private void smiGetByBalls_Click(object sender, EventArgs e)
        {
            FBSClass.MakeFBS(1);
        }

        private void smiEGEStatistics_Click(object sender, EventArgs e)
        {
            //EGE Stat
            new EgeStatistics().Show();
        }
        private void smiEnterMarks_Click(object sender, EventArgs e)
        {
            new SelectVed().Show();
        }
        private void smiLoadMarks_Click(object sender, EventArgs e)
        {
            new SelectVedForLoad(false).Show();
        }        

        private void smiEnterManual_Click(object sender, EventArgs e)
        {
            new SelectExamManual().Show();
        }
        private void smiForm2_Click(object sender, EventArgs e)
        {
            new Form2().Show();
        }

        private void smiAppeal_Click(object sender, EventArgs e)
        {
            new SelectVedForLoad(true).Show();
        }
        private void smiDecryptor_Click(object sender, EventArgs e)
        {
            new Decriptor().Show();
        }
        private void smiEgeLoad_Click(object sender, EventArgs e)
        {
            new LoadEgeMarks().Show();
        }
        private void smiRatingList_Click(object sender, EventArgs e)
        {
            new RatingList(false).Show();
        }                          
        private void smiAbitFacultyIntesection_Click(object sender, EventArgs e)
        {
            new AbitFacultyIntersection().Show();
        }
        private void smiRegionAbitsStat_Click(object sender, EventArgs e)
        {
            new RegionAbitStatistics().Show();
        }
        private void smiRegionStatMarks_Click(object sender, EventArgs e)
        {
            new AbitEgeMarksStatistics().Show();
        }
        private void smiRatingListPasha_Click(object sender, EventArgs e)
        {
            new RatingList(true).Show();
        }
        private void smiGetByFIOPasp2_Click(object sender, EventArgs e)
        {
            FBSClass.MakeFBS(3);
        }

        private void smiRegionFacultyAbitCount_Click(object sender, EventArgs e)
        {
            new RegionFacultyAbitCountStatistics().Show();
        }

        private void smiEntryView_Click(object sender, EventArgs e)
        {
            new EntryViewList().Show();
        }

        private void smiDisEntryView_Click(object sender, EventArgs e)
        {
            new DisEntryViewList().Show();
        }

        private void smiOrderNumbers_Click(object sender, EventArgs e)
        {
            new CardOrderNumbers().Show();
        }

        private void smiRatingBackUp_Click(object sender, EventArgs e)
        {
            new BackUpFix().Show();
        }

        private void smiMakeBackDoc_Click(object sender, EventArgs e)
        {
            SomeMethodsClass.SetBackDocForBudgetInEntryView();
        }

        private void smiDeleteDog_Click(object sender, EventArgs e)
        {
            SomeMethodsClass.DeleteDogFromFirstWave();
        }

        private void smiVTB_Click(object sender, EventArgs e)
        {
            ExportClass.ExportVTB();
        }

        private void smiSberbank_Click(object sender, EventArgs e)
        {
            ExportClass.ExportSber();
        }

        private void номераЗачетокToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExportClass.SetStudyNumbers();
        }

        private void smiExportStudent_Click(object sender, EventArgs e)
        {
            new Migrator().Show();
        }

        private void smiOlympSubjectByRegion_Click(object sender, EventArgs e)
        {
            new OlympSubjectByRegion().Show();
        }

        private void smiOlympRegionBySubject_Click(object sender, EventArgs e)
        {
            new OlympRegionBySubject().Show();
        }

        private void smiOlympAbitBallsAndRatings_Click(object sender, EventArgs e)
        {
            new OlympAbitBallsAndRatings().Show();
        }

        private void региональноToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new OlympLevelAbitRating().Show();
        }

        private void smiRegionAbitStat_Rev_Click(object sender, EventArgs e)
        {
            new RegionAbitStatstics_Reversed().Show();
        }

        private void smiRegionAbitEGEMarksStatistics_Click(object sender, EventArgs e)
        {
           new RegionAbitEGEMarksStatistics().Show();

            // тестовая запись
           MessageBox.Show("");
        }

        //private void smiFISEGE_Migrator_Click(object sender, EventArgs e)
        //{
        //    //new MigratorFISEGE().Show();
        //}

        private void smiOlympCheckList_Click(object sender, EventArgs e)
        {
            new OlympCheckList().Show();
        }

        private void smiExamsResultCSV_Click(object sender, EventArgs e)
        {

        }

        private void smiPriemResults_AbitExamResults_Click(object sender, EventArgs e)
        {
            new FormPriemResults_AbitExamResults().Show();
        }

        private void smiFormB_Click(object sender, EventArgs e)
        {
            new FormB().Show();
        }

        private void smiSplitEntryViews_Click(object sender, EventArgs e)
        {
            SomeMethodsClass.SplitEntryViews();
        }

        private void smiQuerier_Click(object sender, EventArgs e)
        {
            new Querier().ShowDialog();
        }

        private void smiPayDataEntry_Click(object sender, EventArgs e)
        {
            new PayDataEntryList().Show();
        }

        private void smiFormV_Click(object sender, EventArgs e)
        {
            new FormV().Show();
        }

        private void smiCompetitiveGroup_Click(object sender, EventArgs e)
        {
            new CompetitiveGroupList().Show();
        }

        private void vuzNamesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new VuzNameList().Show();
        }

        private void smiLicenseProgram_Click(object sender, EventArgs e)
        {
            new LicenseProgramList().Show();
        }

        private void smiObrazProgram_Click(object sender, EventArgs e)
        {
            new ObrazProgramList().Show();
        }

        private void olympBookToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new OlympBookList().Show();
        }

        private void smiOlympName_Click(object sender, EventArgs e)
        {
            new OlympNameList().Show();
        }

        private void smiOlympSubject_Click(object sender, EventArgs e)
        {
            new OlympSubjectList().Show();
        }

        private void smiProfile_Click(object sender, EventArgs e)
        {
            //new Profile
        }

        private void smiAbitLogChangesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new ListAbitLogChanges().Show();
        }

        private void smiListAbitWithInnerPriorities_Click(object sender, EventArgs e)
        {
            new ListAbitWithInnerPriorities().Show();
        }

        private void smiMyList_Click(object sender, EventArgs e)
        {
            new MyList().Show();
        }

        private void smiRatingWithEGE_Click(object sender, EventArgs e)
        {
            new ListRatingWithEGE().Show();
        }

        private void smiStatGSGU_Click(object sender, EventArgs e)
        {
            new StatFormGSGU().Show();
        }

        private void smiStatGSGUForm2_Click(object sender, EventArgs e)
        {
            new StatFormGSGUForm2().Show();
        }

        private void smiStatGSGUForm1A_Click(object sender, EventArgs e)
        {
            new StatFormGSGUForm1A().Show();
        }

        private void smiRegions_Click(object sender, EventArgs e)
        {

        }

        private void smiCountries_Click(object sender, EventArgs e)
        {

        }

        private void smiLoadExamsResultsToParentExamTool_Click(object sender, EventArgs e)
        {

        }

        private void smiAbitRatingKofGroupChanging_Click(object sender, EventArgs e)
        {
            new ListAbitRatingKofGroupChanging().Show();
        }

        private void smiDisEntryOrderFromReEnter_Click(object sender, EventArgs e)
        {
            new DisEntryFromReEnterViewList().Show();
        }

        private void markToHistoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new ExamsVedMarkToHistory().Show();
        }
    }
}