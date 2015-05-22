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
        Brush BrushText, BrushBackGroundLP, BrushBackGroundOP;
        Color BrushBackGroundOPP, BrushBackGroundProf;

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
            this.Width = 840;
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
                    rbColor_CheckedChanged(this.rbNoColor, null);
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
            cbObrazProgram.SelectedIndexChanged += new EventHandler(cbObrazProgram_SelectedIndexChanged);
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
        void cbObrazProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            //UpdateDataGrid();
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
        private void FillGrid(string abitFilters)
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
            clm.ColumnName = "PersonNum";
            examTable.Columns.Add(clm);

            clm = new DataColumn();
            clm.ColumnName = "ФИО";
            examTable.Columns.Add(clm);

            /*
            clm = new DataColumn();
            clm.ColumnName = "Рег_номер";
            examTable.Columns.Add(clm);
             */

            NewWatch wc = new NewWatch();
            wc.Show();
            wc.SetText("Получение данных по учебным планам...");
///////////////////////////////
            ///// Поиск по Направлениям в QEntry
            string query = @"Select distinct qEntry.LicenseProgramId, qEntry.LicenseProgramName
                                from ed.qEntry " + abitFilters;
            DataTable tbl = _bdc.GetDataSet(query).Tables[0];
            string index = "";
            
            foreach (DataRow rwEntry in tbl.Rows)
            {
                ///////////////////////////////
                ///// Напечатано ли название направления
                bool LicenseNamePrint = false;
                ///////////////////////////////
                ///// Поиск ОБРАЗОВАТЕЛЬНЫХ ПРОГРАММ 
                query = @"Select distinct qEntry.ObrazProgramId, qEntry.ObrazProgramName
                                from ed.qEntry " + abitFilters + " and LicenseProgramId=" + rwEntry.Field<int>("LicenseProgramId").ToString();
                DataTable tbl_LicProg = _bdc.GetDataSet(query).Tables[0];

                bool ObrazProgramNameIsPrinted = false;
                foreach (DataRow rw_licProg in tbl_LicProg.Rows)
                {
                    ///////////////////////////////
                    ///// ДЛЯ КАЖДОЙ ОБРАЗОВАТЕЛЬНОЙ ПРОГРАММЫ ПОИСК ПРОФИЛЕЙ:
                    query = @"select distinct ProfileId, ProfileName from ed.qEntry" + abitFilters + " and LicenseProgramId=" + rwEntry.Field<int>("LicenseProgramId").ToString() +
                            " and ObrazProgramId=" + rw_licProg.Field<int>("ObrazProgramId").ToString() + " and ProfileId is not null";
                    DataTable tbl_ObrProgramProfile = _bdc.GetDataSet(query).Tables[0];
                    ///////////////////////////////
                    /////  ЕСЛИ ЕСТЬ НЕНУЛЕВЫЕ ПРОФИЛИ (ПРОБЛЕМА С ИД СТОЛБЦА)
                    ///// НЕ ДОЛЖНО БЫТЬ ЗАГОЛОВКА СЛОБЦА, СТОЛБЕЦ = (НАПР/ОБРПРОГ/ПРОФ)
                    if (tbl_ObrProgramProfile.Rows.Count > 0)
                    {
                        foreach (DataRow row_profile in tbl_ObrProgramProfile.Rows)
                        {
                            clm = new DataColumn();
                            index = rwEntry.Field<int>("LicenseProgramId").ToString() + "_" + rw_licProg.Field<int>("ObrazProgramId").ToString() + "_" + row_profile.Field<Guid>("ProfileId").ToString();
                            clm.ColumnName = index;
                            examTable.Columns.Add(clm);

                            if (LicenseNamePrint)
                            {
                                row_LicProg[index] = "";
                            }
                            else
                            {
                                row_LicProg[index] = rwEntry.Field<string>("LicenseProgramName"); LicenseNamePrint = true;
                            }
                            if (ObrazProgramNameIsPrinted)
                                row_ObrazProg[index] = "";
                            else
                                row_ObrazProg[index] = rw_licProg.Field<String>("ObrazProgramName");
                            row_Profile[index] = row_profile.Field<string>("ProfileName");
                        }   
                    }
                    ///////////////////////////////
                    /////  НЕНУЛЕВЫХ ПРОФИЛЕЙ НЕТ (ВОЗМОЖНО ЕСТЬ OBRAZ_PROGRAM_IN_ENTRY) 
                    else {
                        ////////////
                        //// нужно получить EntryId 
                        query = @"select distinct qEntry.Id from ed.qEntry" + abitFilters + " and LicenseProgramId=" + rwEntry.Field<int>("LicenseProgramId").ToString() +
                            " and ObrazProgramId=" + rw_licProg.Field<int>("ObrazProgramId").ToString(); 
                        Guid EntryId = (Guid) _bdc.GetDataSet(query).Tables[0].Rows[0].Field<Guid>("Id");
                        ///////////
                        /// поиск по EntryId В ОБРАЗОВАТЕЛЬНЫХ ПРОГРАММАХ
                        query = @"SELECT distinct ObrazProgramInEntry.[Id] as Id, SP_ObrazProgram.Name as Name, SP_ObrazProgram.Id as ObrazProgramId
                              FROM [ed].[ObrazProgramInEntry] 
                             inner join ed.SP_ObrazProgram on ObrazProgramInEntry.ObrazProgramId = SP_ObrazProgram.Id where EntryId ='" + EntryId + @"'
                               order by ObrazProgramId";
                        DataTable tbl_ObrProgram = _bdc.GetDataSet(query).Tables[0];
                        ///////////////////////////////
                        ///// приоритетов образ.программ нет
                        if (tbl_ObrProgram.Rows.Count == 0)
                        {
                            index = rwEntry.Field<int>("LicenseProgramId").ToString() + "_" + rw_licProg.Field<int>("ObrazProgramId").ToString() + "_0";
                            clm = new DataColumn();
                            clm.ColumnName = index;
                            examTable.Columns.Add(clm);
                            
                            if (LicenseNamePrint)
                            {
                                row_LicProg[index] = "";
                            }
                            else
                            {
                                row_LicProg[index] = rwEntry.Field<string>("LicenseProgramName"); LicenseNamePrint = true;
                            }

                            row_ObrazProg[index] = rw_licProg.Field<String>("ObrazProgramName");

                            row_Profile[index] = "(приоритет программы)";
                        }
                        else
                        {
                            ///////////////////////////////
                            ///// ПРИОРИТЕТЫ ОБРАЗ.ПРОГРАММ есть
                            foreach (DataRow rw_ObProg in tbl_ObrProgram.Rows)
                            {
                                clm = new DataColumn();
                                index = rwEntry.Field<int>("LicenseProgramId").ToString() + "_" + rw_ObProg.Field<int>("ObrazProgramId").ToString() + "_0";
                                clm.ColumnName = index;
                                examTable.Columns.Add(clm);
                                row_ObrazProg[index] = rw_ObProg.Field<String>("Name");
                                if (LicenseNamePrint)
                                {
                                    row_LicProg[index] = "";
                                }
                                else
                                {
                                    row_LicProg[index] = rwEntry.Field<string>("LicenseProgramName"); LicenseNamePrint = true;
                                }
                                row_Profile[index] = "(приоритет программы)";

                                ///////////////////////////////
                                ///// ЕСТЬ ЛИ ПРИОРИТЕТНОСТЬ ПРОФИЛЕЙ?
                                query = @"select distinct ProfileInObrazProgramInEntry.Id as Id, SP_Profile.Name as Name, SP_Profile.Id as ProfileId
                                from ed.ProfileInObrazProgramInEntry 
                                inner join ed.SP_Profile on SP_Profile.Id = ProfileInObrazProgramInEntry.ProfileId 
                                where 
                                ObrazProgramInEntryId ='" + rw_ObProg.Field<Guid>("Id") + @"' 
                                order by ProfileId";
                                DataTable tbl_Prof = _bdc.GetDataSet(query).Tables[0];
                                if (tbl_Prof.Rows.Count > 1)
                                {
                                    foreach (DataRow rw_Profile in tbl_Prof.Rows)
                                    {
                                        clm = new DataColumn();
                                        index = rwEntry.Field<int>("LicenseProgramId").ToString() + "_" + rw_ObProg.Field<int>("ObrazProgramId").ToString() + "_" + rw_Profile.Field<int>("ProfileId").ToString();
                                        clm.ColumnName = index;
                                        examTable.Columns.Add(clm);
                                        row_LicProg[index] = "";
                                        row_ObrazProg[index] = "";
                                        row_Profile[index] = rw_Profile.Field<String>("Name");
                                    }
                                }
                                else
                                    if (tbl_Prof.Rows.Count == 1)
                                    {
                                        DataRow rw_Profile = tbl_Prof.Rows[0];
                                        row_Profile[index] = rw_Profile.Field<string>("Name");
                                        index = rwEntry.Field<int>("LicenseProgramId").ToString() + "_" + rw_ObProg.Field<int>("ObrazProgramId").ToString() + "_" + rw_Profile.Field<int>("ProfileId").ToString();
                                        HideColumnsList.Add(index);
                                    }
                            }
                        }
                    }
                }
                // ЗАКОНЧИЛСЯ ПОИСК ВНУТРИ ОБРАЗОВАТЕЛЬНОЙ ПРОГРАММЫ
            }
            examTable.Rows.Add(row_LicProg);
            examTable.Rows.Add(row_ObrazProg);
            examTable.Rows.Add(row_Profile);

            wc.SetText("Получение данных по абитуриентам...(0)");

            int itopList=0;
            if (!String.IsNullOrEmpty(tbAbitsTop.Text))
                if (!int.TryParse(tbAbitsTop.Text, out itopList))
                {
                    itopList = 0;
                }
            string toplist = (rbAbitsAll.Checked)? "": ((itopList == 0)?"":" top "+itopList.ToString());
            query = @"SELECT distinct" + toplist + @"
                        Abiturient.PersonId,
                        extPerson.PersonNum,
                        --Person.Surname +' '+ Person.Name + (case when (Person.SecondName is not null) then ' '+Person.SecondName else '' end) as FIO
                        extPerson.FIO as FIO
                        FROM [ed].Abiturient
                        inner join ed.extPerson on extPerson.Id = Abiturient.PersonId
                        inner join ed.qEntry on qEntry.Id = Abiturient.EntryId 
                        " + abitFilters + " AND Abiturient.Id NOT IN (SELECT Id FROM ed.qAbiturientForeignApplicationsOnly) AND Abiturient.BackDoc = 0"+
                        @" order by FIO";
            tbl = _bdc.GetDataSet(query).Tables[0];
            int personcount = 0;
            int personcountAll = tbl.Rows.Count;
            wc.SetMax(personcountAll);
            foreach (DataRow rw in tbl.Rows)
            {
                wc.PerformStep();
                wc.SetText("Получение данных по абитуриентам...("+personcount+"/"+personcountAll+")");
                personcount++;
                query = @"select 
                                Abiturient.Id, 
                                qEntry.LicenseProgramId, 
                                qEntry.ObrazProgramId as qEntryObrazProgramId ,
                                qEntry.ProfileId as qEntryProfileId,
                                Abiturient.Priority,

                                ObrazProgramInEntry.ObrazProgramId as ObrazProgramId, 
                                ApplicationDetails.ObrazProgramInEntryPriority as ObrazProgramPriority, 
                                ProfileInObrazProgramInEntry.ProfileId as ProfileId,
                                ApplicationDetails.ProfileInObrazProgramInEntryPriority as ProfilePriority

                                from ed.Abiturient

                                inner join ed.qEntry on Abiturient.EntryId = qEntry.Id
                                left join ed.ObrazProgramInEntry on ObrazProgramInEntry.EntryId = qEntry.Id
                                left join ed.ProfileInObrazProgramInEntry on ProfileInObrazProgramInEntry.ObrazProgramInEntryId = ObrazProgramInEntry.Id
                                left join ed.ApplicationDetails on (ApplicationDetails.ApplicationId = Abiturient.Id and (ObrazProgramInEntry.Id = ApplicationDetails.ObrazProgramInEntryId or ObrazProgramInEntry.ObrazProgramId is null) and (ProfileInObrazProgramInEntry.Id = ApplicationDetails.ProfileInObrazProgramInEntryId or ProfileInObrazProgramInEntry.ProfileId is null ))
                                " + abitFilters + " AND Abiturient.BackDoc = 0 and Abiturient.PersonId='" + rw.Field<Guid>("PersonId") + "' " +
                                @"";
                DataRow newRow;
                newRow = examTable.NewRow();
                newRow["Id"] = rw.Field<Guid>("PersonId");
                newRow["PersonNum"] = rw.Field<string>("PersonNum");
                newRow["ФИО"] = rw.Field<string>("FIO");

                DataTable tbl_abit = _bdc.GetDataSet(query).Tables[0];
                foreach (DataRow rw_abit in tbl_abit.Rows)
                {
                    int LicProgId = rw_abit.Field<int>("LicenseProgramId");
                    int EntryObrazProgId = rw_abit.Field<int>("qEntryObrazProgramId");
                    int ObrazProgramId = rw_abit.Field<int?>("ObrazProgramId")??0;
                    string qEntryProfileId = rw_abit.Field<Guid?>("qEntryProfileId").ToString() ?? "";

                    int ProfileId = rw_abit.Field<int?>("ProfileId")??0;

                    int AbitPrior = rw_abit.Field<int?>("Priority") ?? 0;
                    int ObrProgPrior = rw_abit.Field<int?>("ObrazProgramPriority")??0;
                    int ProfPrior = rw_abit.Field<int?>("ProfilePriority") ?? 0;

                    string rowname = "";
                    int Priority = 0;
                    if (ObrazProgramId == 0)
                    {
                        Priority = AbitPrior;
                        if (String.IsNullOrEmpty(qEntryProfileId))
                        {
                            rowname = LicProgId.ToString() + '_' + EntryObrazProgId.ToString() + "_0";
                        }
                        else
                        {
                            rowname = LicProgId.ToString() + '_' + EntryObrazProgId.ToString() + "_" + qEntryProfileId;
                        }
                        newRow[rowname] = Priority == 0 ? "нет" : Priority.ToString();
                    }
                    else
                    {
                        Priority = ObrProgPrior;
                        rowname = LicProgId.ToString() + '_' + ObrazProgramId.ToString() + "_0";
                        newRow[rowname] = Priority == 0 ? "нет" : Priority.ToString();
                        if (!(ProfileId == 0))
                        {
                            Priority = ProfPrior;
                            rowname = LicProgId.ToString() + '_' + ObrazProgramId.ToString() + "_" + ProfileId.ToString();
                            if (!HideColumnsList.Contains(rowname))
                                newRow[rowname] = Priority == 0 ? "нет" : Priority.ToString();
                        }
                    }
                }
                examTable.Rows.Add(newRow);
            } 

            DataView dv = new DataView(examTable);
            dgvAbitList.DataSource = dv;
            dgvAbitList.Columns["Id"].Visible = false;
            wc.Close();
            dgvAbitList.Update();
        }
        private string GetAbitFilterString()
        {
            string s = " WHERE 1=1 ";

            s += " AND ed.qEntry.StudyLevelGroupId IN (" + Util.BuildStringWithCollection(MainClass.lstStudyLevelGroupId) + ")";

            //обработали форму обучения  
            if (StudyFormId != null)
                s += " AND ed.qEntry.StudyFormId = " + StudyFormId;

            //обработали основу обучения  
            if (StudyBasisId != null)
                s += " AND ed.qEntry.StudyBasisId = " + StudyBasisId;

            //обработали факультет
            if (FacultyId != null)
                s += " AND ed.qEntry.FacultyId = " + FacultyId;

            //обработали тип конкурса          
            /*if (CompetitionId != null)
                s += " AND ed.extAbit.CompetitionId = " + CompetitionId;
            */
            //обработали Направление
            if (LicenseProgramId != null)
                s += " AND ed.qEntry.LicenseProgramId = " + LicenseProgramId;

            //обработали Образ программу
            if (ObrazProgramId != null)
                s += " AND ed.qEntry.ObrazProgramId = " + ObrazProgramId;

            //обработали специализацию 
            /*if (ProfileId != null)
                s += string.Format(" AND ed.extAbit.ProfileId = '{0}'", ProfileId);
            */
            return s;
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
            if (dgvAbitList.Rows.Count > 3)
            {
                dgvAbitList.Rows[0].MinimumHeight = 40;
                dgvAbitList.Rows[1].MinimumHeight = 40;
                dgvAbitList.Rows[2].MinimumHeight = 40;
                //foreach (DataGridViewCell cell in dgvAbitList.Rows[2].Cells)
                //{
                //    if (cell.Value.ToString().Contains("приоритет программы"))
                //    {
                //        cell.Style.BackColor = BrushBackGroundOPP;
                //    }
                //    else
                //        if (!String.IsNullOrEmpty(cell.Value.ToString()))
                //        {
                //            cell.Style.BackColor = BrushBackGroundProf;
                //        }
                //}
                if (cbNoPriority.Checked)
                {
                    if (e.ColumnIndex < 2 || dgvAbitList.Rows[e.RowIndex].DefaultCellStyle.BackColor == Color.Red)
                        return;
                    if (dgvAbitList[e.ColumnIndex, e.RowIndex].Value.ToString().Equals("нет"))
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
        }

        private void PaintRectangle(PaintEventArgs e, List<Rectangle> RectangleList, Brush BrushBackground, Brush brushTEXT, string text)
        {
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
            e.Graphics.DrawString(text, this.dgvAbitList.Font, brushTEXT, rect0, stringFormat);
        }

        private void dgvAbitList_Paint(object sender, PaintEventArgs e)
        {
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

            } 
        }

        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if ((DataView)dgvAbitList.DataSource != null)
                PrintToExcel(((DataView)dgvAbitList.DataSource).Table.Copy(), "export");
        }

        private void btnRePaint_Click(object sender, EventArgs e)
        {
            FillGrid(GetAbitFilterString());

            lblCount.Text = "Всего: " + (Dgv.Rows.Count-3).ToString();
            btnCard.Enabled = !(Dgv.RowCount == 0);
        }

        private void PrintToExcel(DataTable tbl, string sheetName)
        {
            if (tbl.Rows.Count <= 3)
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
                    for (int rowindex = 0; rowindex < 2; rowindex++)
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

                    Excel.Range Range3 = ws.Range[ws.Cells[3, 1], ws.Cells[3, tbl.Columns.Count]];
                    Range3.WrapText = true;
                    Range3.RowHeight = rowHeight;
                    Range3.ColumnWidth = colWidth;

                    Range3 = ws.Range[ws.Cells[3, 1], ws.Cells[3, 1]];
                    Range3.ColumnWidth = colNum; 


                    Range3 = ws.Range[ws.Cells[3, 2], ws.Cells[3, 2]];
                    Range3.ColumnWidth = colFIOWidth;

                    Range3 = ws.Range[ws.Cells[1, 1], ws.Cells[tbl.Rows.Count, 2]];
                    Range3.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                    Range3 = ws.Range[ws.Cells[1, 3], ws.Cells[tbl.Rows.Count, tbl.Columns.Count]];
                    Range3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    Range3.NumberFormat = "General"; 


                    for (int rowindex = 2; rowindex < tbl.Rows.Count; rowindex++ )
                    //foreach (DataRow dr in tbl.Rows)
                    {
                        DataRow dr = tbl.Rows[rowindex];
                        j = 1;
                        for (int colindex = 0; colindex<tbl.Columns.Count; colindex++)
                        //foreach (DataColumn dc in tbl.Columns)
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

        private void rbColor_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton rb = sender as RadioButton; 
            if (rb == null)
            {
                MessageBox.Show("Sender is not a RadioButton");
                return;
            }
            if (rb == this.rbNoColor)
            {
                BrushText = DefaultBrushNoColorText;
                BrushBackGroundLP = DefaultBrushNoColorBackGround;
                BrushBackGroundOP = DefaultBrushNoColorBackGround;
                BrushBackGroundOPP = DefaultColorNoColorBackGround;
                BrushBackGroundProf = DefaultColorNoColorBackGround;
            }
            else
                if (rb == this.rbWithColor)
                {
                    BrushText = DefaultBrushWithColorText;
                    BrushBackGroundLP = DefaultBrushWithColorBackGroundLicenseProgram;
                    BrushBackGroundOP = DefaultBrushWithColorBackGroundObrazProgram;
                    BrushBackGroundOPP = DefaultBrushWithColorBackGroundObrazProgramPrior;
                    BrushBackGroundProf = DefaultBrushWithColorBackGroundProfile;
                }
        }

        private void rbWithColor_Click(object sender, EventArgs e)
        {
        }

        private void dgvAbitList_Scroll(object sender, ScrollEventArgs e)
        {
        }

        private void cbNoPriority_CheckStateChanged(object sender, EventArgs e)
        {
            if (dgvAbitList.Rows.Count > 3)
            {
                if (cbNoPriority.Checked)
                {
                    foreach (DataGridViewRow rw in dgvAbitList.Rows)
                    {
                        if (dgvAbitList.Rows.IndexOf(rw) < 3)
                            continue;
                        foreach (DataGridViewCell cell in rw.Cells)
                        {
                            if (cell.ColumnIndex < 2)
                                continue;
                            if (cell.Value.ToString().Equals("нет"))
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
                        if (dgvAbitList.Rows.IndexOf(rw) < 3)
                            continue;
                        rw.DefaultCellStyle.BackColor = Color.White;
                    }
                }
            }
        }

    }
}
