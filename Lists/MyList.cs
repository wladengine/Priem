﻿using System;
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

namespace Priem
{
    public partial class MyList : BookList
    {
        List<Guid> PersonList = new List<Guid>();
        List<string> PersonListFio = new List<string>();
        List<List<KeyValuePair<int, int>>> Coord = new List<List<KeyValuePair<int, int>>>();
        List<List<KeyValuePair<int, int>>> Coord_Save = new List<List<KeyValuePair<int, int>>>();
        List<KeyValuePair<int, KeyValuePair<int, int>>> DeleteList = new List<KeyValuePair<int, KeyValuePair<int, int>>>();
        Guid ErrorGuid = Guid.Empty;

        public MyList()
        {
            InitializeComponent();
            Dgv = dgvAbitList;
            InitControls();
        }

        protected override void ExtraInit()
        {
            base.ExtraInit();
            btnRemove.Visible = btnAdd.Visible = false;
            tbAbitsTop.Text = "10";
            rbAbitsTop.Checked = true;
            _title = "Рейтинговый список с внутренними приоритетами";
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    ComboServ.FillCombo(cbStudyBasis, HelpClass.GetComboListByTable("ed.StudyBasis", "ORDER BY Name"), false, false);
                    FillFaculty();
                    cbStudyBasis.SelectedIndex = 0;
                    FillStudyForm();
                    FillLicenseProgram();
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

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.FacultyId.ToString(), u.FacultyName)).OrderBy(x => x.Value).Distinct().ToList();

                ComboServ.FillCombo(cbFaculty, lst, false, true);
                cbFaculty.SelectedIndex = 0;
            }
        }
        private void FillStudyForm()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = MainClass.GetEntry(context);
                    
                if (FacultyId.HasValue)
                    ent = ent.Where(c => c.FacultyId == FacultyId);

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.StudyFormId.ToString(), u.StudyFormName)).Distinct().OrderBy(u => u.Key).ToList();

                ComboServ.FillCombo(cbStudyForm, lst, false, false);
                cbStudyForm.SelectedIndex = 0;
            }
        }
        private void FillLicenseProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = MainClass.GetEntry(context);

                if (FacultyId.HasValue)
                    ent = ent.Where(c => c.FacultyId == FacultyId);
                if (StudyFormId != null)
                    ent = ent.Where(c => c.StudyFormId == StudyFormId);

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.LicenseProgramId.ToString(), u.LicenseProgramName)).Distinct().ToList();

                ComboServ.FillCombo(cbLicenseProgram, lst, false, true);
                cbLicenseProgram.SelectedIndex = 0;
            }
        }

        public override void UpdateDataGrid()
        {
        }
        private void FillGrid(string abitFilters)
        {
            DataTable examTable = new DataTable();
            examTable.Columns.Add("Id");

            DataRow row_LicProg = examTable.NewRow();
            DataRow row_ObrazProg = examTable.NewRow();
            DataRow row_EntryId = examTable.NewRow();
            DataRow row_ObrazProgramInEntryId = examTable.NewRow();
            DataRow row_KCP = examTable.NewRow();

            DataColumn clm;

            PersonList = new List<Guid>();
            PersonListFio = new List<string>();
            Coord = new List<List<KeyValuePair<int, int>>>();

            NewWatch wc = new NewWatch();
            wc.Show();
            
            ///// Поиск по Направлениям в QEntry
            string query = @"Select distinct qEntry.LicenseProgramId, qEntry.LicenseProgramName
                                from ed.qEntry " + abitFilters;
            DataTable tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            string index = "";
            int cnt = 1;
            foreach (DataRow rwEntry in tbl.Rows)
            {
                wc.SetText("Получение данных по учебным планам... (Обработано конкурсов: " + (cnt++).ToString() + "/" + tbl.Rows.Count + ")");
                ///// Поиск ОБРАЗОВАТЕЛЬНЫХ ПРОГРАММ 
                query = @"Select distinct qEntry.ObrazProgramId, qEntry.ObrazProgramName
                                from ed.qEntry " + abitFilters + " and LicenseProgramId=" + rwEntry.Field<int>("LicenseProgramId").ToString();
                DataTable tbl_LicProg = MainClass.Bdc.GetDataSet(query).Tables[0];

                foreach (DataRow rw_licProg in tbl_LicProg.Rows)
                {
                    ///// ДЛЯ КАЖДОЙ ОБРАЗОВАТЕЛЬНОЙ ПРОГРАММЫ ПОИСК ПРОФИЛЕЙ:
                    query = @"select distinct qEntry.Id, KCP, ProfileId, ProfileName from ed.qEntry" + abitFilters + " and LicenseProgramId=" + rwEntry.Field<int>("LicenseProgramId").ToString() +
                            " and ObrazProgramId=" + rw_licProg.Field<int>("ObrazProgramId").ToString() + " and ProfileId is not null";
                    DataTable tbl_ObrProgramProfile = MainClass.Bdc.GetDataSet(query).Tables[0];
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
                            row_LicProg[index] = rwEntry.Field<string>("LicenseProgramName");
                            row_ObrazProg[index] = rw_licProg.Field<string>("ObrazProgramName");
                            row_EntryId[index] = row_profile.Field<Guid>("Id");
                            row_ObrazProgramInEntryId[index] = "";
                            row_KCP[index] = row_profile.Field<int>("KCP");
                        }
                    }
                    /////  НЕНУЛЕВЫХ ПРОФИЛЕЙ НЕТ (ВОЗМОЖНО ЕСТЬ OBRAZ_PROGRAM_IN_ENTRY) 
                    else
                    {
                        //// нужно получить EntryId 
                        query = @"select distinct qEntry.Id, KCP from ed.qEntry" + abitFilters + " and LicenseProgramId=" + rwEntry.Field<int>("LicenseProgramId").ToString() +
                            " and ObrazProgramId=" + rw_licProg.Field<int>("ObrazProgramId").ToString();
                        DataSet ds = MainClass.Bdc.GetDataSet(query);
                        Guid EntryId = (Guid)ds.Tables[0].Rows[0].Field<Guid>("Id");
                        int _KCP = (int)ds.Tables[0].Rows[0].Field<int>("KCP");

                        /// поиск по EntryId В ОБРАЗОВАТЕЛЬНЫХ ПРОГРАММАХ
                        query = @"SELECT distinct ObrazProgramInEntry.[Id] as Id, SP_ObrazProgram.Name as Name, SP_ObrazProgram.Id as ObrazProgramId, KCP
                              FROM [ed].[ObrazProgramInEntry] 
                             inner join ed.SP_ObrazProgram on ObrazProgramInEntry.ObrazProgramId = SP_ObrazProgram.Id where EntryId ='" + EntryId + @"'
                               order by ObrazProgramId";
                        DataTable tbl_ObrProgram = MainClass.Bdc.GetDataSet(query).Tables[0];
                        ///// приоритетов образ.программ нет
                        if (tbl_ObrProgram.Rows.Count == 0)
                        {
                            index = rwEntry.Field<int>("LicenseProgramId").ToString() + "_" + rw_licProg.Field<int>("ObrazProgramId").ToString() + "_0";
                            clm = new DataColumn();
                            clm.ColumnName = index;
                            examTable.Columns.Add(clm);
                            row_LicProg[index] = rwEntry.Field<string>("LicenseProgramName");
                            row_ObrazProg[index] = rw_licProg.Field<String>("ObrazProgramName");
                            row_EntryId[index] = EntryId.ToString();
                            row_ObrazProgramInEntryId[index] = "";
                            //
                            row_KCP[index] = _KCP;
                        }
                        else
                        {
                            ///// ПРИОРИТЕТЫ ОБРАЗ.ПРОГРАММ есть
                            foreach (DataRow rw_ObProg in tbl_ObrProgram.Rows)
                            {
                                clm = new DataColumn();
                                index = rwEntry.Field<int>("LicenseProgramId").ToString() + "_" + rw_ObProg.Field<int>("ObrazProgramId").ToString() + "_0";
                                clm.ColumnName = index;
                                examTable.Columns.Add(clm);
                                row_ObrazProg[index] = rw_ObProg.Field<String>("Name");
                                row_LicProg[index] = rwEntry.Field<string>("LicenseProgramName");
                                row_EntryId[index] = EntryId;
                                row_ObrazProgramInEntryId[index] = rw_ObProg.Field<Guid>("Id").ToString();
                                row_KCP[index] = rw_ObProg.Field<int>("KCP");
                            }
                        }
                    }
                }
                // ЗАКОНЧИЛСЯ ПОИСК ВНУТРИ ОБРАЗОВАТЕЛЬНОЙ ПРОГРАММЫ
            }

            examTable.Rows.Add(row_LicProg);
            examTable.Rows.Add(row_ObrazProg);
            examTable.Rows.Add(row_EntryId);
            examTable.Rows.Add(row_ObrazProgramInEntryId);
            examTable.Rows.Add(row_KCP);

            wc.SetText("Получение данных по абитуриентам...(0/0)");

            int itopList = 0;
            if (!String.IsNullOrEmpty(tbAbitsTop.Text))
                if (!int.TryParse(tbAbitsTop.Text, out itopList))
                {
                    itopList = 0;
                }
            string toplist = (rbAbitsAll.Checked) ? "" : ((itopList == 0) ? "" : " top " + itopList.ToString());
            /*
             * надо бы вытащить для каждого Entry (cтолбца) вытащить столбец данных с Abiturient учитывая ранжирования по баллу
             */
            List<DataRow> RowList = new List<DataRow>();
            query = @"select " + toplist + @" Abiturient.Id, PersonId, Priority, extAbitMarksSum.TotalSum, (Person.Surname+' '+Person.Name+(case when SecondName is not null then ' '+SecondName else '' end)) as FIO
            from ed.Abiturient
            inner join ed.Person on Abiturient.PersonId = Person.Id
            left join ed.extAbitMarksSum on extAbitMarksSum.Id = Abiturient.Id
            inner join ed.qEntry on Abiturient.EntryId = qEntry.Id
            where Abiturient.EntryId=@EntryId and Abiturient.BackDoc = 0  and Abiturient.IsGosLine=0 
            order by extAbitMarksSum.TotalSum desc";
            wc.SetMax(examTable.Columns.Count-1);
            wc.SetText("Получение данных по абитуриентам...(0/"+(examTable.Columns.Count-1).ToString()+")");
            for (int i = 1; i < examTable.Columns.Count; i++)
            {
                int j = 0;
                string _obrazentryId = examTable.Rows[3][i].ToString();
                String EntryId = examTable.Rows[2][i].ToString();
                DataSet ds = MainClass.Bdc.GetDataSet(query, new SortedList<string, object> { { "@EntryId", EntryId } });
                foreach (DataRow rw in ds.Tables[0].Rows)
                {
                    Guid _AbitId = rw.Field<Guid>("Id");
                    Guid _PersonId = rw.Field<Guid>("PersonId");
                    string FIO = rw.Field<string>("FIO");
                    if (!PersonList.Contains(_PersonId))
                    {
                        PersonList.Add(_PersonId);
                        PersonListFio.Add(FIO);
                    }

                    int _Priority = rw.Field<int?>("Priority") ?? 0;
                    int _obrazPrior = 0;
                    if (!String.IsNullOrEmpty(_obrazentryId))
                    {
                        string query_obrazProgram = @"select ObrazProgramInEntryPriority from ed.ApplicationDetails 
                              where ApplicationDetails.ApplicationId='" + _AbitId.ToString() + "' and ObrazProgramInEntryId='" + _obrazentryId + "'";
                        DataSet dsobrprog = MainClass.Bdc.GetDataSet(query_obrazProgram);
                        if (dsobrprog.Tables[0].Rows.Count > 0)
                            _obrazPrior = dsobrprog.Tables[0].Rows[0].Field<int?>("ObrazProgramInEntryPriority") ?? 0;

                    }
                    String Temp_String = _Priority.ToString() + "_" + _obrazPrior.ToString() + "_" + _PersonId.ToString();
                    Temp_String += "_" + rw.Field<int?>("TotalSum").ToString();
                    if (j < RowList.Count)
                    {
                        DataRow rowTable = RowList[j];
                        rowTable[examTable.Columns[i]] = Temp_String;
                    }
                    else
                    {
                        DataRow rowTable = examTable.NewRow();
                        rowTable[examTable.Columns[i]] = Temp_String;
                        RowList.Add(rowTable);
                    }
                    int tempindex = PersonList.IndexOf(_PersonId);
                    if (Coord.Count <= tempindex)
                        Coord.Add(new List<KeyValuePair<int, int>>());
                    // сначала столбец, потом строка
                    Coord[tempindex].Add(new KeyValuePair<int, int>(i, j));
                    j++;
                }
                wc.PerformStep();
                wc.SetText("Получение данных по абитуриентам...(Обработано конкурсов: "+i+"/" + (examTable.Columns.Count-1).ToString() + ")");
            }
            for (int j = 0; j < RowList.Count; j++)
            {
                DataRow rw = RowList[j];
                examTable.Rows.Add(rw);
            }
            Coord_Save = Coord;

            DataView dv = new DataView(examTable);
            dgvAbitList.DataSource = dv;
            dgvAbitList.Columns["Id"].Visible = false;
            dgvAbitList.AllowUserToOrderColumns = false;
            for (int i = 0; i < dgvAbitList.Columns.Count; i++)
                dgvAbitList.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            wc.Close();
            dgvAbitList.Update();

            PaintGrid();
        }
        private string GetAbitFilterString()
        {
            string s = " WHERE 1=1 ";
            s += " AND ed.qEntry.StudyLevelGroupId = " + MainClass.studyLevelGroupId;

            //обработали форму обучения  
            if (StudyFormId != null)
                s += " AND ed.qEntry.StudyFormId = " + StudyFormId;

            //обработали основу обучения  
            if (StudyBasisId != null)
                s += " AND ed.qEntry.StudyBasisId = " + StudyBasisId;

            //обработали факультет
            if (FacultyId != null)
                s += " AND ed.qEntry.FacultyId = " + FacultyId;
           
            //обработали Направление
            if (LicenseProgramId != null)
                s += " AND ed.qEntry.LicenseProgramId = " + LicenseProgramId;

            return s;
        }
        private void btnFillGrid_Click(object sender, EventArgs e)
        {
            FillGrid(GetAbitFilterString());
        }
        private void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillStudyForm();
        }
        private void cbStudyBasis_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillStudyForm();
        }
        private void cbStudyForm_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillLicenseProgram();
        }
        private void cbLicenseProgram_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private bool UpdateResult()
        {
            bool result = false;
            foreach (DataGridViewRow dw in dgvAbitList.Rows)
            {
                foreach (DataGridViewCell dcell in dw.Cells)
                {
                    if (dcell.Style.BackColor == Color.LightGreen)
                    {
                        string PersId = dcell.Value.ToString();
                        PersId = PersId.Substring(PersId.IndexOf("_") + 1);
                        PersId = PersId.Substring(0, PersId.LastIndexOf("_"));
                        PersId = PersId.Substring(PersId.IndexOf("_") + 1);
                        int index = PersonList.IndexOf(Guid.Parse(PersId));
                        int count = 0;
                        foreach (KeyValuePair<int, int> kvp in Coord[index])
                        {
                            if (dgvAbitList.Rows[kvp.Value + 5].Cells[kvp.Key].Style.BackColor == Color.LightGreen)
                            {
                                count++;
                                if (count > 1)
                                {
                                    ErrorGuid = Guid.Parse(PersId);
                                    return true;
                                }
                            }
                        }
                    }
                }
            }
            return result;
        }

        private void PaintGrid()
        {
            int startrow = 5;
            for (int colindex = 1; colindex < dgvAbitList.Columns.Count; colindex++)
            {
                int KCP = 0;
                int.TryParse(dgvAbitList.Rows[4].Cells[colindex].Value.ToString(), out KCP);

                for (int j = startrow; (j < KCP + startrow) && (j < dgvAbitList.Rows.Count); j++)
                {
                    dgvAbitList.Rows[j].Cells[colindex].Style.BackColor = Color.LightGreen;
                }
            }
            for (int colindex = 1; colindex < dgvAbitList.Columns.Count; colindex++)
            {
                bool hasinnerprior = !String.IsNullOrEmpty((String)dgvAbitList.Rows[3].Cells[colindex].Value);
                for (int j = startrow; j < dgvAbitList.Rows.Count; j++)
                {
                    if (dgvAbitList.Rows[j].Cells[colindex].Value.ToString().StartsWith("0_"))
                    {
                        dgvAbitList.Rows[j].Cells[colindex].Style.BackColor = Color.LightBlue;
                    }
                    if (hasinnerprior)
                        if (dgvAbitList.Rows[j].Cells[colindex].Value.ToString().Contains("_0_"))
                        {
                            dgvAbitList.Rows[j].Cells[colindex].Style.BackColor = Color.LightBlue;
                        }
                }
            }
            dgvAbitList.Update();
            int prior = 0;
            int innerprior = 0;
            int _step = 0;
            NewWatch wc = new NewWatch();
            wc.Show();
            wc.SetText("Анализируем приоритеты ...");
            wc.SetMax(dgvAbitList.Rows.Count);
            while (UpdateResult() || (_step == 0))
            {
                _step++;
                wc.PerformStep();
                if (_step > dgvAbitList.Rows.Count)
                {
                    MessageBox.Show("Цикл перерасчета приоритетов был произведен "+dgvAbitList.Rows.Count +" раз, наверно возникла какая-то ошибка. Придется прекратить перерасчет приоритетов. Проблемный Guid: "+ErrorGuid.ToString(), "Вы знаете,..", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                }
                for (int colindex = 1; colindex < dgvAbitList.Columns.Count; colindex++)
                {
                    int KCP = 0;
                    if (int.TryParse(dgvAbitList.Rows[4].Cells[colindex].Value.ToString(), out KCP))
                    { }


                    bool hasinnerprior = !String.IsNullOrEmpty((String)dgvAbitList.Rows[3].Cells[colindex].Value);
                    for (int j = startrow; (j < dgvAbitList.Rows.Count); j++)
                    {
                        if ((dgvAbitList.Rows[j].Cells[colindex].Style.BackColor != Color.LightGreen) &&
                            (dgvAbitList.Rows[j].Cells[colindex].Style.BackColor != Color.Yellow) &&
                            (dgvAbitList.Rows[j].Cells[colindex].Style.BackColor != Color.LightBlue))
                        {
                            break;
                        }
                        if (dgvAbitList.Rows[j].Cells[colindex].Style.BackColor == Color.LightGreen)
                        {
                            string cellvalue = dgvAbitList.Rows[j].Cells[colindex].Value.ToString();
                            // пока только первый приоритет
                            string temp = cellvalue.Substring(0, cellvalue.IndexOf('_'));
                            if (!int.TryParse(temp, out prior))
                            {
                                dgvAbitList.Rows[j].Cells[colindex].Style.BackColor = Color.Red;
                            }
                            cellvalue = cellvalue.Substring(cellvalue.IndexOf('_') + 1);
                            // внутренний приоритет
                            temp = cellvalue.Substring(0, cellvalue.IndexOf('_'));
                            if (!int.TryParse(temp, out innerprior))
                            {
                                dgvAbitList.Rows[j].Cells[colindex].Style.BackColor = Color.Red;
                            }
                            // получили PersonId 
                            cellvalue = cellvalue.Substring(cellvalue.IndexOf('_') + 1);
                            cellvalue = cellvalue.Substring(0, cellvalue.IndexOf('_'));
                            // пора обновить грид, нашли в списке PersonId
                            int index = PersonList.IndexOf(Guid.Parse(cellvalue));
                            // если он был:
                            if (index > -1)
                            {
                                if (cellvalue.StartsWith("c6a065"))
                                { }
                                // по всем координатам key;value = столбец; строка
                                foreach (KeyValuePair<int, int> kvp in Coord[index])
                                {

                                    // если это та же строка и тот же столбец
                                    if ((kvp.Value + startrow == j) && (kvp.Key == colindex))
                                    {
                                        continue;
                                    }

                                    int KCP_temp = 0;
                                    if (int.TryParse(dgvAbitList.Rows[4].Cells[kvp.Key].Value.ToString(), out KCP_temp))
                                    { }

                                    cellvalue = dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Value.ToString();
                                    temp = cellvalue.Substring(0, cellvalue.IndexOf('_'));
                                    int prior_temp = 0;
                                    if (!int.TryParse(temp, out prior_temp))
                                    {
                                        dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Style.BackColor = Color.Red;
                                        continue;
                                    }
                                    if ((prior_temp == prior) && (hasinnerprior))
                                    {
                                        cellvalue = cellvalue.Substring(cellvalue.IndexOf('_') + 1);
                                        // внутренний приоритет
                                        int innerprior_temp = 0;
                                        temp = cellvalue.Substring(0, cellvalue.IndexOf('_'));
                                        if (!int.TryParse(temp, out innerprior_temp))
                                        {
                                            dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Style.BackColor = Color.Red;
                                            continue;
                                        }
                                        if (innerprior_temp > innerprior)
                                        {
                                            bool isGreen = false;
                                            if (dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Style.BackColor == Color.LightGreen)
                                            {
                                                isGreen = true;
                                            }
                                            if (kvp.Key == 1)
                                                if (kvp.Value == 18)
                                                { }
                                            dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Style.BackColor = Color.Yellow;

                                            DeleteList.Add(new KeyValuePair<int, KeyValuePair<int, int>>(index, kvp));
                                            if (isGreen)
                                            {
                                                for (int row_temp = startrow + KCP_temp; row_temp < dgvAbitList.Rows.Count; row_temp++)
                                                {
                                                    if ((dgvAbitList.Rows[row_temp].Cells[kvp.Key].Style.BackColor != Color.LightGreen) &&
                                                   (dgvAbitList.Rows[row_temp].Cells[kvp.Key].Style.BackColor != Color.Yellow) &&
                                                       (dgvAbitList.Rows[row_temp].Cells[kvp.Key].Style.BackColor != Color.LightBlue))
                                                    //if (dgvAbitList.Rows[row_temp].Cells[kvp.Key].Style.BackColor == Color.White)
                                                    {
                                                        dgvAbitList.Rows[row_temp].Cells[kvp.Key].Style.BackColor = Color.LightGreen;
                                                        break;
                                                    }
                                                }
                                            }
                                        }

                                    }
                                    else
                                    {
                                        if (prior_temp > prior)
                                        {
                                            bool isGreen = false;
                                            if (dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Style.BackColor == Color.LightGreen)
                                            {
                                                isGreen = true;
                                            }
                                            dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Style.BackColor = Color.Yellow;
                                            //Coord[index].Remove(kvp);
                                            DeleteList.Add(new KeyValuePair<int, KeyValuePair<int, int>>(index, kvp));
                                            if (isGreen)
                                            {
                                                for (int row_temp = startrow + KCP_temp; row_temp < dgvAbitList.Rows.Count; row_temp++)
                                                {
                                                    if ((dgvAbitList.Rows[row_temp].Cells[kvp.Key].Style.BackColor != Color.LightGreen) &&
                                                    (dgvAbitList.Rows[row_temp].Cells[kvp.Key].Style.BackColor != Color.Yellow) &&
                                                        (dgvAbitList.Rows[row_temp].Cells[kvp.Key].Style.BackColor != Color.LightBlue))
                                                    //if (dgvAbitList.Rows[row_temp].Cells[kvp.Key].Style.BackColor == Color.White)
                                                    {
                                                        dgvAbitList.Rows[row_temp].Cells[kvp.Key].Style.BackColor = Color.LightGreen;
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                foreach (KeyValuePair<int, KeyValuePair<int, int>> kvp in DeleteList)
                                {
                                    Coord[kvp.Key].Remove(kvp.Value);
                                }
                                DeleteList = new List<KeyValuePair<int, KeyValuePair<int, int>>>();
                            }
                        }
                    }
                }
            }
            wc.Close();
            CopyTable();
        }
        private void CopyTable()
        {
            for (int i = 5; i< dgvAbitList.Rows.Count; i++)
            {
                foreach (DataGridViewCell dcell in dgvAbitList.Rows[i].Cells)
                {
                    if (String.IsNullOrEmpty(dcell.Value.ToString()))
                        continue;
                    string dcell_Value = dcell.Value.ToString();
                    // приоритет
                    string _prior = dcell_Value.Substring(0, dcell_Value.IndexOf('_'));
                    dcell_Value = dcell_Value.Substring(dcell_Value.IndexOf('_') + 1);
                    // внутренний приоритет
                    string _prior_inner = dcell_Value.Substring(0, dcell_Value.IndexOf('_'));
                    dcell_Value = dcell_Value.Substring(dcell_Value.IndexOf('_') + 1);
                    // Фамилия
                    string FIO = dcell_Value.Substring(0, dcell_Value.IndexOf('_'));
                    FIO = PersonListFio[PersonList.IndexOf(Guid.Parse(FIO))];
                    string ball = dcell_Value.Substring(dcell_Value.IndexOf('_') + 1);
                    dcell.Value = FIO +" ("+_prior+", "+_prior_inner+", "+ball+")";
                }
            }
            dgvAbitList.Rows[2].Visible = false;
            dgvAbitList.Rows[3].Visible = false;
        }

        private void dgvAbitList_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 5)
                return;
            if (e.ColumnIndex < 1)
                return;

            if (dgvAbitList.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor == Color.LightGreen)
                return;

            string FIO = dgvAbitList.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Substring(0, dgvAbitList.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().IndexOf('(')-1);
            int index = PersonListFio.IndexOf(FIO);
            if (index>-1)
                foreach (KeyValuePair<int, int> kvp in Coord_Save[index])
                {
                    if (dgvAbitList.Rows[kvp.Value + 5].Cells[kvp.Key].Style.BackColor == Color.LightGreen)
                    {
                        dgvAbitList.CurrentCell = dgvAbitList.Rows[kvp.Value + 5].Cells[kvp.Key];
                    }
                }
        }

        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if (dgvAbitList.Rows.Count>5)
            {
                byte[][] tableColors = new byte[dgvAbitList.Rows.Count][];
                for (int i = 0; i < tableColors.Length; i++)
                {
                    tableColors[i] = new byte[dgvAbitList.Columns.Count];
                    for (int j = 0; j < tableColors[i].Length; j++)
                    {
                        if (dgvAbitList.Rows[i].Cells[j].Style.BackColor == Color.LightGreen)
                            tableColors[i][j] = 1;
                        else
                            if (dgvAbitList.Rows[i].Cells[j].Style.BackColor == Color.Yellow)
                                tableColors[i][j] = 2;
                            else
                                if (dgvAbitList.Rows[i].Cells[j].Style.BackColor == Color.LightBlue)
                                    tableColors[i][j] = 3;
                                else
                                    tableColors[i][j] = 0;

                    }
                }

                DataTable tbl = ((DataView)dgvAbitList.DataSource).Table;

                string sheetName = "export";

                

                if (tbl.Columns.Contains("Id"))
                {
                    tbl.Columns.Remove("Id");
                }
                for (int i = 0; i < tableColors.Length; i++)
                {
                    for (int j = 1; j < tableColors[i].Length; j++)
                    {
                        if (tableColors[i][j] == 0)
                            continue;
                        if (tableColors[i][j] == 1)
                            dgvAbitList.Rows[i].Cells[j-1].Style.BackColor = Color.LightGreen;
                        else
                            if (tableColors[i][j] == 2)
                                dgvAbitList.Rows[i].Cells[j-1].Style.BackColor = Color.Yellow;
                            else
                                if (tableColors[i][j] == 3)
                                    dgvAbitList.Rows[i].Cells[j-1].Style.BackColor = Color.LightBlue;
                    }
                }
                tbl.Rows[3].Delete();
                tbl.Rows[2].Delete();


                

                List<string> lstFields = new List<string>();

                int rowHeight = 70;
                double colFIOWidth = 50;

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

                        int i = 1;
                        int j = 1;

                        ProgressForm prog = new ProgressForm(0, tbl.Rows.Count, 1, ProgressBarStyle.Blocks, "Импорт списка");
                        prog.Show();

                        // печать из грида

                        //печать строк 0 и 1 - направление и образовательная программа

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
                                    j_end_merge = colindex + 1;
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
                                    j_begin_merge = colindex + 1;
                                    j_end_merge = colindex + 1;
                                    ws.Cells[i, colindex + 1] = dr[dc.ColumnName] == null ? "" : "'" + text;
                                    Excel.Range Range0 = ws.Range[ws.Cells[i, colindex + 1], ws.Cells[i, colindex + 1]];
                                    Range0.WrapText = true;
                                    if (i > 1)
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
                        // тут был раньше профиль. (Если убрать Entry, то тут должно быть КСР)
                        Excel.Range Range3 = ws.Range[ws.Cells[1, 1], ws.Cells[3, tbl.Columns.Count]];
                        Range3.WrapText = true;
                        Range3.RowHeight = rowHeight;
                        Range3.ColumnWidth = colFIOWidth;
                        Range3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        Range3.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                        Range3 = ws.Range[ws.Cells[4, 1], ws.Cells[tbl.Rows.Count, tbl.Columns.Count]];
                        Range3.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                        // начиная со второй строки (КСР, и далее абитуриенты)
                        for (int rowindex = 2; rowindex < tbl.Rows.Count; rowindex++)
                        {
                            DataRow dr = tbl.Rows[rowindex];
                            j = 1;
                            for (int colindex = 0; colindex < tbl.Columns.Count; colindex++)
                            {
                                DataColumn dc = tbl.Columns[colindex];
                                ws.Cells[i, j] = dr[dc.ColumnName] == null ? "" : dr[dc.ColumnName].ToString();
                                Range3 = ws.Cells[i, j];
                                if (tableColors[rowindex+2][colindex+1] !=0)
                                    Range3.Interior.Color = dgvAbitList.Rows[rowindex].Cells[colindex].Style.BackColor;
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
        }
    }
}