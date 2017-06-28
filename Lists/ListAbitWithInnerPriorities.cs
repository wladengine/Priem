using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

using WordOut;
using EducServLib;
using BDClassLib;
using BaseFormsLib;
using PriemLib;

namespace Priem
{
    public partial class ListAbitWithInnerPriorities : BookList
    {
        List<string> HideColumnsList;
        int LastSystemRowIndex;

        Brush DefaultBrushNoColorText = Brushes.Black;
        Brush DefaultBrushNoColorBackGround = Brushes.White;
        Color DefaultColorNoColorBackGround = Color.White;

        Brush DefaultBrushWithColorText = Brushes.Black;
        Brush DefaultBrushWithColorBackGroundLicenseProgram = Brushes.White;
        Brush DefaultBrushWithColorBackGroundObrazProgram = Brushes.White;
        Color DefaultBrushWithColorBackGroundObrazProgramPrior = Color.White;
        Color DefaultBrushWithColorBackGroundProfile = Color.White;

        public ListAbitWithInnerPriorities()
        {
            InitializeComponent();
            this._title = "Список конкурсов. Внутренние приоритеты";
            Dgv = dgvAbitList; 
            InitControls();
        }
        //дополнительная инициализация контролов
        protected override void ExtraInit()
        {
            base.ExtraInit();
            lblCount.Text = "";
            btnRemove.Visible = btnAdd.Visible = false;
            HideColumnsList = new List<string>();
            tbAbitsTop.Text = "100";
            rbAbitsAll.Checked = true;
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    ComboServ.FillCombo(cbStudyBasis, HelpClass.GetComboListByTable("ed.StudyBasis", "ORDER BY Name"), false, false);
                    FillFaculty();
                    cbStudyBasis.SelectedIndex = 0;
                    FillStudyForm();
                    FillLicenseProgram();
                    FillObrazProgram();
                    //MessageBox.Show("Для отображения результатов, выберите необходимые значения полей и нажмите кнопку 'Обновить данные'", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка при инициализации формы " + exc.Message);
            }
        }
        #region Handlers

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
        #endregion

        private void FillFaculty()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = MainClass.GetEntry(context);

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.FacultyId.ToString(), u.FacultyName)).Distinct().ToList();

                ComboServ.FillCombo(cbFaculty, lst, false, false);
                cbFaculty.SelectedIndex = 0;
            }
        }
        private void FillStudyForm()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = MainClass.GetEntry(context).Where(c => c.FacultyId == FacultyId);

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.StudyFormId.ToString(), u.StudyFormName)).Distinct().OrderBy(u => u.Key).ToList();

                ComboServ.FillCombo(cbStudyForm, lst, false, false);
                cbStudyForm.SelectedIndex = 0;
            }
        }
        private void FillLicenseProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = MainClass.GetEntry(context).Where(c => c.FacultyId == FacultyId);

                if (StudyFormId != null)
                    ent = ent.Where(c => c.StudyFormId == StudyFormId);

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.LicenseProgramId.ToString(), u.LicenseProgramName)).Distinct().ToList();

                ComboServ.FillCombo(cbLicenseProgram, lst, false, true);
                cbLicenseProgram.SelectedIndex = 0;
            }
        }
        private void FillObrazProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = MainClass.GetEntry(context).Where(c => c.FacultyId == FacultyId);

                if (StudyFormId != null)
                    ent = ent.Where(c => c.StudyFormId == StudyFormId);
                if (LicenseProgramId != null)
                    ent = ent.Where(c => c.LicenseProgramId == LicenseProgramId);

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.ObrazProgramId.ToString(), u.ObrazProgramName + ' ' + u.ObrazProgramCrypt)).Distinct().ToList();

                ComboServ.FillCombo(cbObrazProgram, lst, false, true);
            }
        }

        public override void InitHandlers()
        {
            cbFaculty.SelectedIndexChanged += new EventHandler(cbFaculty_SelectedIndexChanged);
            cbLicenseProgram.SelectedIndexChanged += new EventHandler(cbLicenseProgram_SelectedIndexChanged);
            cbStudyForm.SelectedIndexChanged += new EventHandler(cbStudyForm_SelectedIndexChanged);
            cbStudyBasis.SelectedIndexChanged += new EventHandler(cbStudyBasis_SelectedIndexChanged);
        }
        void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {  
            FillStudyForm(); 
        }
        void cbStudyForm_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillLicenseProgram();
        }
        void cbLicenseProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillObrazProgram();
        }
        void cbStudyBasis_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillLicenseProgram();
        }

        public override void UpdateDataGrid()
        {
            /*
            FillGrid(GetAbitFilterString());

            lblCount.Text = "Всего: " + Dgv.Rows.Count.ToString();
            btnCard.Enabled = !(Dgv.RowCount == 0);
             */
        }
        private void FillGrid()
        {
            using (PriemEntities context = new PriemEntities())
            {
                DataTable examTable = new DataTable();

                DataRow row_LicProg = examTable.NewRow();
                DataRow row_ObrazProg = examTable.NewRow();
                DataRow row_Profile = examTable.NewRow();

                DataColumn clm;

                clm = new DataColumn();
                clm.ColumnName = "Id";
                examTable.Columns.Add(clm);

                clm = new DataColumn();
                clm.ColumnName = "Number";
                examTable.Columns.Add(clm);

                clm = new DataColumn();
                clm.ColumnName = "PersonNum";
                examTable.Columns.Add(clm);

                clm = new DataColumn();
                clm.ColumnName = "ФИО";
                examTable.Columns.Add(clm);

                clm = new DataColumn();
                clm.ColumnName = "Оригиналы";
                examTable.Columns.Add(clm);

                NewWatch wc = new NewWatch();
                wc.Show();
                wc.SetText("Получение данных по учебным планам...");
                ///////////////////////////////

                var Entry_List = (from entry in context.qEntry
                                  join in_en in context.InnerEntryInEntry on entry.Id equals in_en.EntryId into gj
                                  from inner_entry in gj.DefaultIfEmpty()
                                  where
                                    (FacultyId.HasValue ? entry.FacultyId == FacultyId : true)
                                  && (MainClass.lstStudyLevelGroupId.Contains(entry.StudyLevelGroupId))
                                  && (LicenseProgramId.HasValue ? entry.LicenseProgramId == LicenseProgramId : true)
                                  && (ObrazProgramId.HasValue ? entry.ObrazProgramId == ObrazProgramId : true)
                                  && (StudyBasisId.HasValue ? entry.StudyBasisId == StudyBasisId : true)
                                  && (StudyFormId.HasValue ? entry.StudyFormId == StudyFormId : true)
                                  && entry.IsForeign == rbIsForeign.Checked
                                  select new
                                  {
                                      entry.Id,
                                      entry.StudyLevelGroupId,
                                      entry.StudyBasisId,
                                      entry.StudyFormId,
                                      entry.LicenseProgramId,
                                      entry.LicenseProgramName,
                                      entry.ObrazProgramId,
                                      entry.ObrazProgramName,
                                      entry.ProfileId,
                                      entry.ProfileName,
                                      HasInnerEntryInEntry = (inner_entry != null),
                                      InnerEntryInEntryId = (inner_entry == null ? Guid.Empty : inner_entry.Id),
                                      InnerObrazProgramId = (inner_entry == null ? -1 : inner_entry.ObrazProgramId),
                                      InnerProfileId = (inner_entry == null ? -1 : inner_entry.ProfileId),
                                  }).OrderBy(x=>x.LicenseProgramName).ToList();

                List<int> StudyFormList = Entry_List.Select(x => x.StudyFormId).Distinct().ToList();
                foreach (int StudyForm_id in StudyFormList)
                {
                    List<int> StudyBasisList = Entry_List.Where(x => x.StudyFormId == StudyForm_id).Select(x => x.StudyBasisId).Distinct().ToList();
                    foreach (int StudyBasis_id in StudyBasisList)
                    {
                        List<int> LicenseProgramList = Entry_List.Where(x => x.StudyFormId == StudyForm_id && x.StudyBasisId == StudyBasis_id).Select(e => e.LicenseProgramId).Distinct().ToList();
                        foreach (int LP_id in LicenseProgramList)
                        {
                            var Tbl_lp = Entry_List.Where(x => x.StudyFormId == StudyForm_id && x.StudyBasisId == StudyBasis_id && x.LicenseProgramId == LP_id).Select(x => x).ToList();
                            string LicProgram_Name = Tbl_lp.Select(x => x.LicenseProgramName).First();
                            List<int> ObrazProgramList = Tbl_lp.Select(x => x.ObrazProgramId).Distinct().ToList();
                            foreach (int ObProgram_id in ObrazProgramList)
                            {
                                List<int> ProfileList = Tbl_lp.Where(x => x.ObrazProgramId == ObProgram_id).Select(x => x.ProfileId).Distinct().ToList();
                                string ObrazProgramName = Tbl_lp.Where(x => x.ObrazProgramId == ObProgram_id).Select(x => x.ObrazProgramName).First().ToString();
                                if (ProfileList.Count == 1 && ProfileList[0] == 0)
                                {
                                    Guid EntryId = Tbl_lp.Where(x => x.ObrazProgramId == ObProgram_id && x.ProfileId == ProfileList[0]).Select(x => x.Id).First();
                                    List<int> InnerObrazProgramList = Tbl_lp.Where(x => x.Id == EntryId && x.HasInnerEntryInEntry).Select(x => x.InnerObrazProgramId).Distinct().ToList();
                                    if (InnerObrazProgramList.Count == 0)
                                    {
                                        clm = new DataColumn();
                                        clm.ColumnName = EntryId.ToString() + "_" + Guid.Empty.ToString();
                                        examTable.Columns.Add(clm);
                                        row_LicProg[clm] = LicProgram_Name;
                                        row_ObrazProg[clm] = ObrazProgramName;
                                        row_Profile[clm] = Tbl_lp.Where(x => x.ObrazProgramId == ObProgram_id && x.ProfileId == ProfileList[0]).Select(x => x.ProfileName).First().ToString();
                                    }
                                    else
                                    {
                                        clm = new DataColumn();
                                        clm.ColumnName = EntryId.ToString() + "_" + Guid.Empty.ToString();
                                        examTable.Columns.Add(clm);
                                        row_LicProg[clm] = LicProgram_Name;
                                        row_ObrazProg[clm] = ObrazProgramName;
                                        row_Profile[clm] = "Приоритет заявления";

                                        foreach (int InnerObrazProgram_id in InnerObrazProgramList)
                                        {
                                            List<int> InnerProfileList = Tbl_lp.Where(x => x.Id == EntryId && x.HasInnerEntryInEntry && x.InnerObrazProgramId == InnerObrazProgram_id).Select(x => x.InnerProfileId).Distinct().ToList();
                                            string InnerObrazProgramName = context.SP_ObrazProgram.Where(x => x.Id == InnerObrazProgram_id).Select(x => x.Name).First();
                                            foreach (int InnerProfile_id in InnerProfileList)
                                            {
                                                Guid InnerEntryId = Tbl_lp.Where(x => x.InnerObrazProgramId == InnerObrazProgram_id && x.InnerProfileId == InnerProfile_id).Select(x => x.InnerEntryInEntryId).First();
                                                string InnerProfileName = context.SP_Profile.Where(x => x.Id == InnerProfile_id).Select(x => x.Name).First();
                                                clm = new DataColumn();
                                                clm.ColumnName = EntryId.ToString() + "_" + InnerEntryId.ToString();
                                                examTable.Columns.Add(clm);
                                                row_LicProg[clm] = LicProgram_Name;
                                                row_ObrazProg[clm] = InnerObrazProgramName;
                                                row_Profile[clm] = InnerProfileName;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    foreach (int Profile_id in ProfileList)
                                    {
                                        Guid EntryId = Tbl_lp.Where(x => x.ObrazProgramId == ObrazProgramId && x.ProfileId == Profile_id).Select(x => x.Id).First();
                                        clm = new DataColumn();
                                        clm.ColumnName = EntryId.ToString() + "_" + Guid.Empty.ToString();
                                        examTable.Columns.Add(clm);
                                        row_LicProg[clm] = LicProgram_Name;
                                        row_ObrazProg[clm] = ObrazProgramName;
                                        row_Profile[clm] = Tbl_lp.Where(x => x.ObrazProgramId == ObrazProgramId && x.ProfileId == ProfileList[0]).Select(x => x.ProfileName).First().ToString();
                                    }
                                }
                            }
                        }
                    }
                }
                row_Profile["Number"] = "№";
                row_Profile["PersonNum"] = "ИД";
                row_Profile["ФИО"] = "ФИО";
                row_Profile["Оригиналы"] = "Оригиналы";
                examTable.Rows.Add(row_LicProg);
                examTable.Rows.Add(row_ObrazProg);
                examTable.Rows.Add(row_Profile);
                LastSystemRowIndex = examTable.Rows.Count;
                wc.SetText("Получение данных по абитуриентам...");

                var Abits = (from abit in context.extAbit
                             join entry in context.qEntry on abit.EntryId equals entry.Id
                             join pers in context.extPerson on abit.PersonId equals pers.Id

                             join in_en in context.InnerEntryInEntry on abit.EntryId equals in_en.EntryId into gjj
                             from inner_entry in gjj.DefaultIfEmpty()

                             join app_in_en in context.ApplicationDetails on 
                                 new { AppId = abit.Id, InnerEntryId =  inner_entry.Id} 
                                 equals new { AppId = app_in_en.ApplicationId, InnerEntryId = app_in_en.InnerEntryInEntryId } into gj
                             from app_inner_entry_prior in gj.DefaultIfEmpty()
                             where
                                    (FacultyId.HasValue ? entry.FacultyId == FacultyId : true)
                                  && (MainClass.lstStudyLevelGroupId.Contains(entry.StudyLevelGroupId))
                                  && (LicenseProgramId.HasValue ? entry.LicenseProgramId == LicenseProgramId : true)
                                  && (ObrazProgramId.HasValue ? entry.ObrazProgramId == ObrazProgramId : true)
                                  && (StudyBasisId.HasValue ? entry.StudyBasisId == StudyBasisId : true)
                                  && (StudyFormId.HasValue ? entry.StudyFormId == StudyFormId : true)
                                  && entry.IsForeign == rbIsForeign.Checked
                             select new
                             {
                                 abit.PersonId, 
                                 abit.FIO,
                                 abit.EntryId,
                                 abit.Priority,
                                 pers.Barcode,
                                 pers.PersonNum,
                                 pers.HasOriginals,
                                 InnerEntryInEntryId = (inner_entry == null ? Guid.Empty : inner_entry.Id),
                                 InnerPrior = (app_inner_entry_prior == null ? -1 : app_inner_entry_prior.InnerEntryInEntryPriority),
                             }).ToList();
                int i = 1;
                List<int?> AbitList = Abits.Select(x => x.Barcode).Distinct().ToList();
                wc.SetText("Получение данных по абитуриентам...0/" + AbitList.Count.ToString());
                wc.SetMax(AbitList.Count);
                foreach (var barcode in AbitList)
                {
                    DataRow rw = examTable.NewRow();
                    rw["Id"] = Abits.Where(x => x.Barcode == barcode).Select(x => x.PersonId).First();
                    rw["Number"] = i.ToString(); 
                    rw["ФИО"] = Abits.Where(x => x.Barcode == barcode).Select(x => x.FIO).First().ToString();
                    rw["PersonNum"] = Abits.Where(x => x.Barcode == barcode).Select(x => x.PersonNum).First().ToString();
                    rw["Оригиналы"] = Abits.Where(x => x.Barcode == barcode).Select(x => x.HasOriginals).First().Value ? "да" : "нет";
                    var Apps = Abits.Where(x => x.Barcode == barcode).Select(x => x).ToList();
                    foreach (var app in Apps)
                    {
                        string colname = app.EntryId.ToString() + "_" + app.InnerEntryInEntryId.ToString();
                        if (examTable.Columns.Contains(colname))
                            rw[colname] = (app.InnerEntryInEntryId == Guid.Empty) ? app.Priority.ToString() : (app.InnerPrior > 0 ? app.InnerPrior.ToString() : "нет приоритета");
                        if (app.InnerEntryInEntryId != Guid.Empty)
                        {
                            rw[app.EntryId.ToString() + "_" + Guid.Empty.ToString()] = app.Priority.ToString();
                        }

                    }
                    examTable.Rows.Add(rw);
                    wc.PerformStep();
                    wc.SetText("Получение данных по абитуриентам..."+i.ToString()+"/" + AbitList.Count.ToString());
                    i++;
                }
                DataView dv = new DataView(examTable);
                dgvAbitList.DataSource = dv;
                dgvAbitList.Columns["Id"].Visible = false;
                wc.Close();
                if (chbWithOriginals.Checked)
                    SetVisibleRows();
                dgvAbitList.Update();

                dgvAbitList.ColumnHeadersVisible = false;
                if (dgvAbitList.Rows.Count > LastSystemRowIndex)
                {
                    dgvAbitList.Rows[0].MinimumHeight = 40;
                    dgvAbitList.Rows[1].MinimumHeight = 40;
                    dgvAbitList.Rows[2].MinimumHeight = 40;
                }
                dgvAbitList.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                dgvAbitList.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                string[] ColumnsName = { "ФИО", "Number", "PersonNum" ,"Оригиналы"};
                foreach (string s in ColumnsName)
                {
                    int indexColumnId = dgvAbitList.Columns[s].Index;
                    if (indexColumnId >= 0)
                    {
                        dgvAbitList.Columns[indexColumnId].DefaultCellStyle.WrapMode = DataGridViewTriState.False;
                        dgvAbitList.Columns[indexColumnId].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        dgvAbitList.Columns[indexColumnId].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    }
                }
                
            }
        }
        protected override void OpenCard(string itemId, BaseFormEx formOwner, int? index)
        {
            MainClass.OpenCardPerson(itemId, this, dgvAbitList.CurrentRow.Index);
        }
        //поиск по фио
        private void tbFIO_TextChanged(object sender, EventArgs e)
        {
            WinFormsServ.Search(this.dgvAbitList, "ФИО", tbFIO.Text);
        }

        private void dgvAbitList_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            dgvAbitList.ColumnHeadersVisible = false;
            if (dgvAbitList.Rows.Count > LastSystemRowIndex)
            {
                dgvAbitList.Rows[0].MinimumHeight = 40;
                dgvAbitList.Rows[1].MinimumHeight = 40;
                dgvAbitList.Rows[2].MinimumHeight = 40;
                
                if (cbNoPriority.Checked && e.RowIndex>LastSystemRowIndex)
                {
                    if (e.ColumnIndex < 2 || dgvAbitList.Rows[e.RowIndex].DefaultCellStyle.BackColor == Color.Red)
                        return;
                    if (dgvAbitList[e.ColumnIndex, e.RowIndex].Value.ToString().Contains("нет"))
                    {
                        dgvAbitList.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Red;
                    }
                }
            }
            dgvAbitList.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvAbitList.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            int indexColumnId = dgvAbitList.Columns["ФИО"].Index;
            if (indexColumnId >= 0)
            {
                dgvAbitList.Columns[indexColumnId].DefaultCellStyle.WrapMode = DataGridViewTriState.False;
                dgvAbitList.Columns[indexColumnId].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                dgvAbitList.Columns[indexColumnId].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

        }

        private void dgvAbitList_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            /*
            if (dgvAbitList.CurrentCell == null)
                return;
            if (dgvAbitList.CurrentCell.RowIndex<0)
                return;

            Guid PersonId = Guid.Parse(dgvAbitList.CurrentRow.Cells["Id"].Value.ToString());
            */
        }

        private void PaintRectangle(PaintEventArgs e, List<Rectangle> RectangleList, Brush BrushBackground, Brush brushTEXT, string text)
        {/*
            Rectangle rect0 = RectangleList[0];
            for (int k = 1; k < RectangleList.Count; k++)
            {
                rect0.Width += RectangleList[k].Width - 1;
            }
            rect0.Width -= 1;
            rect0.Height -= 1;
            e.Graphics.FillRectangle(BrushBackground, rect0);
            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;
            stringFormat.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString(text, this.dgvAbitList.Font, brushTEXT, rect0, stringFormat);*/
        }

        private void dgvAbitList_Paint(object sender, PaintEventArgs e)
        {
            /*
            int firstcolumn = 1;
            if (dgvAbitList.Columns.Count > 3)
            for (int i = 0; i < 2; i++)
            {
                DataGridViewRow rw = dgvAbitList.Rows[i];
                List<Rectangle> RectangleList = new List<Rectangle>();
                Rectangle t = this.dgvAbitList.GetCellDisplayRectangle(firstcolumn, i, true);
                string text = rw.Cells[firstcolumn].Value.ToString();
                RectangleList.Add(t);

                for (int j = firstcolumn+1; j < dgvAbitList.Columns.Count; j++)
                {
                    DataGridViewCell cell = dgvAbitList.Rows[i].Cells[j];
                    if (cell.Value.ToString().Equals(""))
                    {
                        Rectangle t1 = this.dgvAbitList.GetCellDisplayRectangle(j, i, true);
                        RectangleList.Add(t1);
                    }
                    else
                    {
                        if (i == 0)
                            PaintRectangle(e, RectangleList, BrushBackGroundLP, BrushText, text);
                        else if (i == 1)
                            PaintRectangle(e, RectangleList, BrushBackGroundOP, BrushText, text);

                        // новый первый в список
                        RectangleList = new List<Rectangle>();
                        t = this.dgvAbitList.GetCellDisplayRectangle(j, i, true);
                        RectangleList.Add(t);
                        text = cell.Value.ToString();
                    }
                }
                if (i == 0)
                    PaintRectangle(e, RectangleList, BrushBackGroundLP, BrushText, text);
                else if (i == 1)
                    PaintRectangle(e, RectangleList, BrushBackGroundOP, BrushText, text);

            } */
        }

        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if ((DataView)dgvAbitList.DataSource != null)
                PrintToExcel(((DataView)dgvAbitList.DataSource).Table.Copy(), "export");
        }

        private void btnRePaint_Click(object sender, EventArgs e)
        {
            FillGrid();
            lblCount.Text = "Всего: " + (from DataGridViewRow rw in dgvAbitList.Rows where rw.Index >= LastSystemRowIndex && rw.Visible select rw).Count().ToString();
            btnCard.Enabled = (Dgv.RowCount > LastSystemRowIndex);
        }

        private void PrintToExcel(DataTable tbl, string sheetName)
        {
            if (tbl.Rows.Count <= LastSystemRowIndex)
                return;

            if (tbl.Columns.Contains("Id"))
            {
                tbl.Columns.Remove("Id");
            }
            
            List<string> lstFields = new List<string>();

            int rowHeight = 70;
            double colWidth = 21;
            double colNum = 10;
            double colFIOWidth = 35;

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Файлы Excel (.xls)|*.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Excel.Application exc = new Excel.Application();
                    Excel.Workbook wb = exc.Workbooks.Add(System.Reflection.Missing.Value);
                    Excel.Worksheet ws = (Excel.Worksheet)exc.ActiveSheet;
                    
                    ws.Name = sheetName.Substring(0, sheetName.Length < 30 ? sheetName.Length - 1 : 30);
                     
                    int i = 0;
                    int j = 1;
                    
                    i++;

                    ProgressForm prog = new ProgressForm(0, tbl.Rows.Count, 1, ProgressBarStyle.Blocks, "Импорт списка");
                    prog.Show();
                    
                    // печать из грида
                    for (int rowindex = 0; rowindex < LastSystemRowIndex; rowindex++)
                    {
                        DataRow dr = tbl.Rows[rowindex];
                        int j_begin_merge = 1;
                        int j_end_merge = 1;
                        for (int colindex = 0; colindex < tbl.Columns.Count; colindex++)
                        {
                            DataColumn dc = tbl.Columns[colindex];
                            string text = dr[dc.ColumnName].ToString();
                            if (String.IsNullOrEmpty(text))
                            {
                                // добавляем к списку или перезаписываем последнюю ячейку
                                j_end_merge = colindex+1;
                            }
                            else
                            {
                                // создаем новый список и объединяем старый
                                if (j_begin_merge != j_end_merge)
                                {
                                    Excel.Range Range1 = ws.Range[ws.Cells[i, j_begin_merge], ws.Cells[i, j_end_merge]];
                                    Range1.Merge();
                                    Range1.WrapText = true;
                                    if (i > 1)
                                        Range1.RowHeight = rowHeight; 
                                    Range1.WrapText = true;
                                }
                                j_begin_merge = colindex+1;
                                j_end_merge = colindex+1;
                                ws.Cells[i, colindex + 1] = dr[dc.ColumnName] == null ? "" : "'" + text;
                                Excel.Range Range0 = ws.Range[ws.Cells[i, colindex + 1], ws.Cells[i, colindex + 1]];
                                Range0.WrapText = true;
                                if (i>1) 
                                    Range0.RowHeight = rowHeight;
                            } 
                        }
                        if (j_begin_merge != j_end_merge)
                        {
                            Excel.Range Range1 = ws.Range[ws.Cells[i, j_begin_merge], ws.Cells[i, j_end_merge]];
                            Range1.Merge();
                            if (i > 1)
                                Range1.RowHeight = rowHeight;
                            Range1.WrapText = true;
                        }
                        i++;
                        prog.PerformStep();
                    }

                    Excel.Range Range3 = ws.Range[ws.Cells[LastSystemRowIndex, 1], ws.Cells[LastSystemRowIndex, tbl.Columns.Count]];
                    Range3.WrapText = true;
                    Range3.RowHeight = rowHeight;
                    Range3.ColumnWidth = colWidth;

                    Range3 = ws.Range[ws.Cells[LastSystemRowIndex, 1], ws.Cells[LastSystemRowIndex, 1]];
                    Range3.ColumnWidth = colNum;

                    Range3 = ws.Range[ws.Cells[LastSystemRowIndex, 2], ws.Cells[LastSystemRowIndex, 2]];
                    Range3.ColumnWidth = colNum;

                    Range3 = ws.Range[ws.Cells[LastSystemRowIndex, 3], ws.Cells[LastSystemRowIndex, 3]];
                    Range3.ColumnWidth = colFIOWidth;

                    Range3 = ws.Range[ws.Cells[1, 1], ws.Cells[tbl.Rows.Count, 3]];
                    Range3.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                    Range3 = ws.Range[ws.Cells[1, 3], ws.Cells[tbl.Rows.Count, tbl.Columns.Count]];
                    Range3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    Range3.NumberFormat = "General";


                    for (int rowindex = LastSystemRowIndex; rowindex < tbl.Rows.Count; rowindex++)
                    //foreach (DataRow dr in tbl.Rows)
                    {
                        DataRow dr = tbl.Rows[rowindex];
                        j = 1;
                        for (int colindex = 0; colindex<tbl.Columns.Count; colindex++)
                        {
                            DataColumn dc = tbl.Columns[colindex];
                            ws.Cells[i, j] = dr[dc.ColumnName] == null ? "" : dr[dc.ColumnName].ToString();
                            j++;
                        }

                        i++;
                        prog.PerformStep();
                    }
                    prog.Close();

                    wb.SaveAs(sfd.FileName, Excel.XlFileFormat.xlExcel7,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        Excel.XlSaveAsAccessMode.xlExclusive,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value);
                    exc.Visible = true;

                }
                catch (System.Runtime.InteropServices.COMException exc)
                {
                    MessageBox.Show(exc.Message);
                }
            }
            //На всякий случай
            sfd.Dispose();
        }

        private void rbWithColor_Click(object sender, EventArgs e)
        {
        }

        private void dgvAbitList_Scroll(object sender, ScrollEventArgs e)
        {
        }

        private void cbNoPriority_CheckedChanged(object sender, EventArgs e)
        {
            if (dgvAbitList.Rows.Count > LastSystemRowIndex)
            {
                if (cbNoPriority.Checked)
                {
                    foreach (DataGridViewRow rw in dgvAbitList.Rows)
                    {
                        if (dgvAbitList.Rows.IndexOf(rw) < LastSystemRowIndex)
                            continue;
                        foreach (DataGridViewCell cell in rw.Cells)
                        {
                            if (cell.ColumnIndex < 2)
                                continue;
                            if (cell.Value.ToString().Contains("нет"))
                            {
                                rw.DefaultCellStyle.BackColor = Color.Red;
                                break;
                            }
                        }
                    }
                }
                else
                {
                    foreach (DataGridViewRow rw in dgvAbitList.Rows)
                    {
                        if (dgvAbitList.Rows.IndexOf(rw) < LastSystemRowIndex)
                            continue;
                        rw.DefaultCellStyle.BackColor = Color.White;
                    }
                }
            }
        }

        private void chbWithOriginals_CheckedChanged(object sender, EventArgs e)
        {
            SetVisibleRows();
        }
        private void SetVisibleRows()
        {
            int Num = 1;
            for (int i = LastSystemRowIndex; i < dgvAbitList.Rows.Count; i++)
            {
                bool isVis = !chbWithOriginals.Checked || (dgvAbitList.Rows[i].Cells["Оригиналы"].Value.ToString() == "да");
                dgvAbitList.Rows[i].Visible = isVis;
                if (isVis)
                    dgvAbitList.Rows[i].Cells["Number"].Value = (Num++).ToString();
            }
            lblCount.Text = "Всего: " + (from DataGridViewRow rw in dgvAbitList.Rows where rw.Index >= LastSystemRowIndex && rw.Visible select rw).Count().ToString();
            btnCard.Enabled = (Num > 1);

        }

    }
}
