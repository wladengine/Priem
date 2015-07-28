using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using EducServLib;
using PriemLib;

namespace Priem
{
    public partial class ListAbitLogChanges : Form
    {
        public int? FacultyId
        {
            get { return ComboServ.GetComboIdInt(cbFaculty); }
            set { ComboServ.SetComboId(cbFaculty, value); }
        }
        public int? LicenseProgramId
        {
            get { return ComboServ.GetComboIdInt(cbLicenseProgram); }
            set { ComboServ.SetComboId(cbLicenseProgram, value); }
        }
        public int? ObrazProgramId
        {
            get { return ComboServ.GetComboIdInt(cbObrazProgram); }
            set { ComboServ.SetComboId(cbObrazProgram, value); }
        }
        public int? StudyBasisId
        {
            get { return ComboServ.GetComboIdInt(cbStudyBasis); }
            set { ComboServ.SetComboId(cbStudyBasis, value); }
        }
        public int? StudyFormId
        {
            get { return ComboServ.GetComboIdInt(cbStudyForm); }
            set { ComboServ.SetComboId(cbStudyForm, value); }
        }

        public BackgroundWorker bw;

        public ListAbitLogChanges()
        {
            InitializeComponent();

            this.MdiParent = MainClass.mainform;

            bw = new BackgroundWorker();
            bw.DoWork += DoBWWork;
            bw.RunWorkerCompleted += bw_RunWorkerCompleted;

            FillComboStudyFormStudyBasis();
            FillComboFaculty();
            FillComboLicenseProgram();
            FillComboObrazProgram();
            FillGrid();
            InitHandlers();
        }

        private void InitHandlers()
        {
            cbFaculty.SelectedIndexChanged += cbFaculty_SelectedIndexChanged;
            cbLicenseProgram.SelectedIndexChanged += cbLicenseProgram_SelectedIndexChanged;
            cbObrazProgram.SelectedIndexChanged += cbObrazProgram_SelectedIndexChanged;
            cbStudyForm.SelectedIndexChanged += cbStudyForm_SelectedIndexChanged;
            cbStudyBasis.SelectedIndexChanged += cbStudyBasis_SelectedIndexChanged;
        }

        private void FillComboStudyFormStudyBasis()
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<KeyValuePair<string, string>> lst = context.StudyForm.Select(x => new { x.Id, x.Name }).Distinct().ToList().
                    Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();

                ComboServ.FillCombo(cbStudyForm, lst, false, true);

                lst = context.StudyBasis.Select(x => new { x.Id, x.Name }).Distinct().ToList().
                    Select(x => new KeyValuePair<string, string>(x.Id.ToString(), x.Name)).ToList();
                ComboServ.FillCombo(cbStudyBasis, lst, false, true);
            }
        }

        private void FillComboFaculty()
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<KeyValuePair<string, string>> lst = context.qEntry.Where(x => MainClass.lstStudyLevelGroupId.Contains(x.StudyLevelGroupId)).Select(x => new { x.FacultyId, x.FacultyName }).Distinct().ToList().
                    Select(x => new KeyValuePair<string, string>(x.FacultyId.ToString(), x.FacultyName)).ToList();

                ComboServ.FillCombo(cbFaculty, lst, false, true);
            }
        }
        private void FillComboLicenseProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<KeyValuePair<string, string>> lst = context.qEntry.Where(x => (FacultyId.HasValue ? x.FacultyId == FacultyId : true) && (MainClass.lstStudyLevelGroupId.Contains(x.StudyLevelGroupId)))
                    .Select(x => new { x.LicenseProgramId, x.LicenseProgramCode, x.LicenseProgramName }).Distinct().ToList().
                    Select(x => new KeyValuePair<string, string>(x.LicenseProgramId.ToString(), x.LicenseProgramCode + " " + x.LicenseProgramName)).ToList();

                ComboServ.FillCombo(cbLicenseProgram, lst, false, true);
            }
        }
        private void FillComboObrazProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<KeyValuePair<string, string>> lst = context.qEntry.Where(x => (FacultyId.HasValue ? x.FacultyId == FacultyId : true)
                    && (LicenseProgramId.HasValue ? x.LicenseProgramId == LicenseProgramId : true) && (MainClass.lstStudyLevelGroupId.Contains(x.StudyLevelGroupId)))
                    .Select(x => new { x.ObrazProgramId, x.ObrazProgramCrypt, x.ObrazProgramName }).Distinct().ToList().
                    Select(x => new KeyValuePair<string, string>(x.ObrazProgramId.ToString(), x.ObrazProgramCrypt + " " + x.ObrazProgramName)).ToList();

                ComboServ.FillCombo(cbObrazProgram, lst, false, true);
            }
        }

        private void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboLicenseProgram();
        }
        private void cbLicenseProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillComboObrazProgram();
        }
        private void cbObrazProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillGrid();
        }
        private void cbStudyForm_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillGrid();
        }
        private void cbStudyBasis_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillGrid();
        }

        private void FillGrid()
        {
            cbFaculty.Enabled = false;
            cbLicenseProgram.Enabled = false;
            cbObrazProgram.Enabled = false;
            cbStudyForm.Enabled = false;
            cbStudyBasis.Enabled = false;
            gbWait.Visible = true;

            bw.RunWorkerAsync(new { FacultyId, LicenseProgramId, ObrazProgramId, StudyFormId, StudyBasisId });

            //dgv.Columns["Id"].Visible = false;
            //dgv.Columns["PersonNum"].HeaderText = "Ид. номер";
            //dgv.Columns["FIO"].HeaderText = "ФИО";
            //dgv.Columns["RegNum"].HeaderText = "Рег. номер";
            //dgv.Columns["LP"].HeaderText = "Направление";
            //dgv.Columns["OP"].HeaderText = "Обр. программа";
            //dgv.Columns["ProfileName"].HeaderText = "Профиль";
            //dgv.Columns["StudyFormName"].HeaderText = "Форма обучения";
            //dgv.Columns["StudyBasisName"].HeaderText = "Основа обучения";
            //dgv.Columns["ActionType"].HeaderText = "Событие";
            //dgv.Columns["ActionTime"].HeaderText = "Время";
            //dgv.Columns["ActionAuthor"].HeaderText = "Автор";
        }

        void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            dgv.DataSource = e.Result;

            dgv.Columns["Id"].Visible = false;
            dgv.Columns["PersonNum"].HeaderText = "Ид. номер";
            dgv.Columns["FIO"].HeaderText = "ФИО";
            dgv.Columns["RegNum"].HeaderText = "Рег. номер";
            dgv.Columns["LP"].HeaderText = "Направление";
            dgv.Columns["OP"].HeaderText = "Обр. программа";
            dgv.Columns["ProfileName"].HeaderText = "Профиль";
            dgv.Columns["StudyFormName"].HeaderText = "Форма обучения";
            dgv.Columns["StudyBasisName"].HeaderText = "Основа обучения";
            dgv.Columns["ActionType"].HeaderText = "Событие";
            dgv.Columns["ActionTime"].HeaderText = "Время";
            dgv.Columns["ActionAuthor"].HeaderText = "Автор";

            cbFaculty.Enabled = true;
            cbLicenseProgram.Enabled = true;
            cbObrazProgram.Enabled = true;
            cbStudyForm.Enabled = true;
            cbStudyBasis.Enabled = true;
            gbWait.Visible = false;
        }

        private void DoBWWork(object data, DoWorkEventArgs args)
        {
            args.Result = GetData(((dynamic)args.Argument).FacultyId, 
                ((dynamic)args.Argument).LicenseProgramId, 
                ((dynamic)args.Argument).ObrazProgramId, 
                ((dynamic)args.Argument).StudyFormId, 
                ((dynamic)args.Argument).StudyBasisId);
        }

        private DataTable GetData(int? iFacultyId, int? iLicenseProgramId, int? iObrazProgramId, int? iStudyFormId, int? iStudyBasisId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                var FacultyIds = context.qFaculty.Select(x => x.Id);
                var data = context.hlpAbiturientActionLog.Where(x => MainClass.lstStudyLevelGroupId.Contains(x.StudyLevelGroupId) && FacultyIds.Contains(x.FacultyId));
                if (iFacultyId.HasValue)
                    data = data.Where(x => x.FacultyId == iFacultyId);
                if (iLicenseProgramId.HasValue)
                    data = data.Where(x => x.LicenseProgramId == iLicenseProgramId);
                if (iObrazProgramId.HasValue)
                    data = data.Where(x => x.ObrazProgramId == iObrazProgramId);
                if (iStudyFormId.HasValue)
                    data = data.Where(x => x.StudyFormId == iStudyFormId);
                if (iStudyBasisId.HasValue)
                    data = data.Where(x => x.StudyBasisId == iStudyBasisId);

                DateTime dtFrom = dtpDateFrom.Value.Date;
                data = data.Where(x => x.ActionTime >= dtFrom);
                DateTime dtTo = dtpDateTo.Value.Date.AddDays(1).AddSeconds(-1);
                data = data.Where(x => x.ActionTime < dtTo);

                var _data = data.Select(x => new
                    {
                        x.Id,
                        x.PersonNum,
                        x.FIO,
                        x.RegNum,
                        LP = x.LicenseProgramCode + " " + x.LicenseProgramName,
                        OP = x.ObrazProgramCrypt + " " + x.ObrazProgramName,
                        x.ProfileName,
                        x.StudyFormName,
                        x.StudyBasisName,
                        x.ActionType,
                        x.ActionTime,
                        x.ActionAuthor
                    }).ToArray();

                return Converter.ConvertToDataTable(_data);
            }
        }

        private void dgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            MainClass.OpenCardAbit(dgv["Id", e.RowIndex].Value.ToString(), null, null);
        }
    }
}
