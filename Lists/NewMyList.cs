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
    public partial class NewMyList : BookList
    {
        List<clEntry> EntryList;
        PersonCoordinates Coord = new PersonCoordinates();
        int LastSystemRow = 0;

        public NewMyList()
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
            btn_GreenList.Visible = false;
            if (MainClass.IsOwner())
            {
                btn_GreenList.Visible = true;
            }
            btn_GreenList.Enabled = false;
            if (MainClass.dbType == PriemType.PriemMag)
            {
                labelWhite.Location = labelThistle.Location;
                pictureBoxWhite.Location = pictureBoxThistle.Location;
                labelBeige.Visible = pictureBoxBeige.Visible = false;
                labelThistle.Visible = false;
                pictureBoxThistle.Visible = false;
            }
            else
            {
                cbZeroWave.Visible = false;
            }
            _title = "Рейтинговый список с внутренними приоритетами";
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    ComboServ.FillCombo(cbStudyBasis, HelpClass.GetComboListByTable("ed.StudyBasis", "ORDER BY Name"), false, true);
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

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.StudyFormId.ToString(), u.StudyFormName)).Distinct().OrderBy(u => u.Key).ToList();

                ComboServ.FillCombo(cbStudyForm, lst, false, true);
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
                else
                {
                    ComboServ.FillCombo(cbLicenseProgram, new List<KeyValuePair<string, string>>(), false, true);
                    return;
                }

                if (StudyFormId != null)
                    ent = ent.Where(c => c.StudyFormId == StudyFormId);
                if (StudyBasisId.HasValue)
                    ent = ent.Where(c => c.StudyBasisId == StudyBasisId);

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
            Coord = new PersonCoordinates();

            NewWatch wc = new NewWatch();
            
            string query = @"Select distinct 
            extEntry.Id as EntryId,
            extEntry.FacultyId, 
            extEntry.StudyBasisId, 
            extEntry.StudyFormId,
            extEntry.LicenseProgramId, 
            extEntry.LicenseProgramName, 
            extEntry.ObrazProgramId, 
            extEntry.ObrazProgramName, 
            extEntry.KCP as extEntryKCP,
            InnerEntryInEntry.Id as InnerEntryInEntryId,
            InnerEntryInEntry.ObrazProgramId,
            SP_ObrazProgram.Name as InnerObrazProgramName,
            InnerEntryInEntry.ProfileId,
            SP_Profile.Name as InnerProfileName,
            InnerEntryInEntry.KCP as InnerEntryInEntryKCP,
            InnerEntryInEntry.EgeExamNameId,
            EgeExamName.Name as EgeName,
            InnerEntryInEntry.EgeMin
            from ed.extEntry 
            left join ed.InnerEntryInEntry on InnerEntryInEntry.EntryId = extEntry.Id
            left join ed.SP_ObrazProgram on InnerEntryInEntry.ObrazProgramId = SP_ObrazProgram.Id
            left join ed.SP_Profile on SP_Profile.Id = InnerEntryInEntry.ProfileId
            left join ed.EgeExamName on EgeExamName.Id = InnerEntryInEntry.EgeExamNameId
            " + abitFilters +
            " order by StudyFormId, StudyBasisId, LicenseProgramName, ObrazProgramName, InnerObrazProgramName, InnerProfileName";

            EntryList = new List<clEntry>();

            DataTable tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            wc.SetMax(1);
            wc.Show();
            List<Guid> entr = (from DataRow rw in tbl.Rows select rw.Field<Guid>("EntryId")).Distinct().ToList();

            foreach (Guid entry in entr)
            {
                wc.SetText("Получение данных по учебным планам... (Обработано конкурсов: " + EntryList.Count.ToString() + "/" + entr.Count + ")");
                var X = (from DataRow rw in tbl.Rows where rw.Field<Guid>("EntryId")==entry select rw).Distinct().ToList();

                clEntry cl = new clEntry();
                cl.EntryId = entry.ToString();
                cl.FacultyId = X[0].Field<int>("FacultyId");
                cl.LicenseProgramName = X[0].Field<string>("LicenseProgramName");
                cl.ObrazProgramName = X[0].Field<string>("ObrazProgramName");
                cl.KCP = X[0].Field<int>("extEntryKCP");
                cl.MaxCountGreen = cl.KCP;
                cl.StudyBasis = (X[0].Field<int>("StudyFormId") == 1) ? "Очная" : "Очно-заочная";
                cl.StudyForm = (X[0].Field<int>("StudyBasisId") == 1) ? "Бюджетная" : "Договорная";

                foreach (var x in X)
                {
                    Column col = new Column();
                    col.InnerEntryId = x.Field<Guid?>("InnerEntryInEntryId").ToString();
                    col.InnerObrazProgramName = String.IsNullOrEmpty(x.Field<string>("InnerObrazProgramName")) ? "(нет)" : x.Field<string>("InnerObrazProgramName");
                    col.MaxKCP = x.Field<int?>("InnerEntryInEntryKCP") ?? x.Field<int>("extEntryKCP");
                    col.ProfileName = String.IsNullOrEmpty(x.Field<string>("InnerProfileName")) ? "(нет)" : x.Field<string>("InnerProfileName");
                    cl.AddColumn(col);
                }
                EntryList.Add(cl);
            }

            int proc = int.Parse(tbDinamicWave.Text);
            double DinamicWave = 1.0 * proc / 100;

            foreach (var x in EntryList)
            {
                wc.SetText("Получение данных по контрольным цифрам приема... (Обработано конкурсов: " + EntryList.IndexOf(x).ToString() + "/" + entr.Count + ")");

                if (!cbZeroWave.Checked)
                {
                    x.KCP = x.KCP - int.Parse(MainClass.Bdc.GetStringValue(@"
                                    select COUNT( distinct extEntryView.Id) 
                                    from ed.extEntryView
                                    inner join ed.Abiturient on AbiturientId = Abiturient.Id
                                    join ed.Entry on Abiturient.EntryId = Entry.Id 
                                    where Abiturient.EntryId = '" + x.EntryId.ToString() + @"'
                                    and Abiturient.CompetitionId NOT IN (12,11)
                                     or 
                                    (Entry.ParentEntryId = extEntryView.EntryId  )
                                    "));
                    foreach (Column col in x.ColumnList)
                    {
                        if (String.IsNullOrEmpty(col.InnerEntryId))
                        { 
                            col.MaxKCP = x.KCP;
                        }
                        else
                        { 
                            col.MaxKCP = col.MaxKCP - int.Parse(MainClass.Bdc.GetStringValue(@"
                                    select COUNT ( distinct extEntryView.Id) 
                                    from ed.extEntryView
                                    inner join ed.Abiturient on AbiturientId = Abiturient.Id
                                    join ed.Entry on extEntryView.EntryId = Entry.Id
                                    join ed.InnerEntryInEntry on InnerEntryInEntry.EntryId = Entry.Id
                                    where Abiturient.EntryId = '" + x.EntryId + @"' and 
                                    Abiturient.InnerEntryInEntryId = '" + col.InnerEntryId + @"'
                                    and Abiturient.CompetitionId NOT IN (12,11)
                                    or 
                                    (Entry.ParentEntryId = extEntryView.EntryId and InnerEntryInEntry.ParentInnerEntryInEntryId = Abiturient.InnerEntryInEntryId)
                                    "));
                        }
                    }
                }
            }
            DataTable examTable = new DataTable();
            DataRow row_LP = examTable.NewRow();
            DataRow row_ObP = examTable.NewRow();
            DataRow row_EntryId = examTable.NewRow();
            DataRow row_innerEntryId = examTable.NewRow();
            DataRow row_inObP = examTable.NewRow();
            DataRow row_KCP = examTable.NewRow();
            foreach (var x in EntryList)
            {
                foreach (Column col in x.ColumnList)
                {
                    DataColumn c = new DataColumn();
                    c.DataType = typeof(bool);
                    examTable.Columns.Add(c);
                    c = new DataColumn();
                    c.ColumnName = col.ColumnName;
                    examTable.Columns.Add(c);
                    col.ColumnIndex = examTable.Columns.Count - 1;
                    row_LP[c] = x.LicenseProgramName;
                    row_ObP[c] = x.ObrazProgramName;
                    row_EntryId[c] = x.EntryId;
                    row_innerEntryId[c] = col.InnerEntryId.ToString();
                    row_inObP[c] = col.InnerObrazProgramName;
                    row_KCP[c] = col.MaxKCP.ToString() + " (" + x.KCP.ToString() + ")";
                }
            }
            examTable.Rows.Add(row_LP);
            examTable.Rows.Add(row_ObP);
            examTable.Rows.Add(row_EntryId);
            examTable.Rows.Add(row_innerEntryId);
            examTable.Rows.Add(row_inObP);
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
             * надо бы вытащить для каждого Entry (cтолбца) вытащить столбец данных с Abiturient учитывая ранжирование, и указывая баллы
             */
            List<DataRow> RowList = new List<DataRow>();
            string Wave = "_FirstWave";
            if (cbZeroWave.Checked)
                Wave = "_ZeroWave";
            query = @"select " + toplist + @" Abiturient.Id, extPerson.PersonNum, extPerson.HasOriginals, 
            Abiturient.PersonId, Abiturient.Priority
            --,  extAbitMarksSum.TotalSum
            , extPerson.FIO as FIO
            from ed.Abiturient
            inner join ed.extPerson on Abiturient.PersonId = extPerson.Id
           -- left join ed.extAbitMarksSum on extAbitMarksSum.Id = Abiturient.Id
            inner join ed." + Wave + @" on " + Wave + @".AbiturientId = Abiturient.Id
            inner join ed.extEntry on Abiturient.EntryId = extEntry.Id
            " + ((cbZeroWave.Checked) ? "inner join ed.extEntryView on extEntryView.AbiturientId = Abiturient.Id" : "") +
            @"
            where Abiturient.EntryId=@EntryId and Abiturient.BackDoc = 0
            and Abiturient.CompetitionId NOT IN (12,11)
            --order by extAbitMarksSum.TotalSum desc
             order by  " + Wave + @".SortNum 
            ";
            wc.SetMax(EntryList.Count);
            foreach (var x in EntryList)
            {
                int entryid = EntryList.IndexOf(x);

                wc.SetText("Получение данных по абитуриентам...("+entryid.ToString()+"/"+EntryList.Count.ToString()+")");
                wc.PerformStep();

                DataSet ds = MainClass.Bdc.GetDataSet(query, new SortedList<string, object> { { "@EntryId", x.EntryId } });
                foreach (DataRow rw in ds.Tables[0].Rows)
                {
                    Abiturient Abit = new Abiturient();
                    Abit.AbitId = rw.Field<Guid>("Id");
                    Abit.PersonId = rw.Field<Guid>("PersonId");
                    Abit.HasOriginals = rw.Field<bool>("HasOriginals");
                    Abit.regNum_FIO = rw.Field<string>("PersonNum") + "_" + rw.Field<string>("FIO");
                    Abit.Priority = rw.Field<int?>("Priority") ?? 0;
                    x.Abits.Add(Abit);
                    Coord.Add(Abit.PersonId, new Coordinates() { entryindex = entryid, abitlistindex = x.Abits.Count - 1, InCompetition = true });
                    if (x.ColumnList.Count > 1)
                    {
                        foreach (var col in x.ColumnList)
                        {
                            int inner = 0;
                            DataSet dsobrprog = MainClass.Bdc.GetDataSet(@"select InnerEntryInEntryPriority from ed.ApplicationDetails 
                              where ApplicationDetails.ApplicationId='" + Abit.AbitId.ToString() + "' and InnerEntryInEntryId='" + col.InnerEntryId + "'");
                            if (dsobrprog.Tables[0].Rows.Count > 0)
                               inner = dsobrprog.Tables[0].Rows[0].Field<int?>("InnerEntryInEntryPriority") ?? 0;
                            col.InnerPriorities.Add(inner);
                            col.AbitColorListAdd();
                        }
                    }
                    else
                    {
                        x.ColumnList[0].InnerPriorities.Add(0);
                        x.ColumnList[0].AbitColorListAdd();
                    }
                }
            }
           
           
            int examTableRowsCountBeforeAbits = examTable.Rows.Count;
            LastSystemRow = examTableRowsCountBeforeAbits;
            foreach (var entry in EntryList)
            {
                int row = 0;
                int i = 0;
                foreach (Abiturient Ab in entry.Abits)
                {
                    if (examTable.Rows.Count == examTableRowsCountBeforeAbits + row)
                    {
                        examTable.Rows.Add(examTable.NewRow());
                    }
                    foreach (Column Col in entry.ColumnList)
                    {
                        examTable.Rows[row + examTableRowsCountBeforeAbits].SetField<bool>(Col.ColumnIndex - 1, Ab.HasOriginals);
                        examTable.Rows[row+examTableRowsCountBeforeAbits].SetField<string>(Col.ColumnName, Ab.regNum_FIO + "(" + Ab.Priority.ToString() + ", " + Col.InnerPriorities[i] + ")");
                    } 
                    i++;
                    row++;
                }
            }
            
            dgvAbitList.DataSource = new DataView(examTable);

            PaintGrid();
            PaintDataGridView(); 

            dgvAbitList.ColumnHeadersVisible = false;
            dgvAbitList.AllowUserToOrderColumns = false;
            for (int i = 0; i < dgvAbitList.Columns.Count; i++)
            {
                dgvAbitList.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                if (i % 2 == 0)
                {
                    dgvAbitList.Columns[i].Width = 20;
                    dgvAbitList.Columns[i].Resizable = DataGridViewTriState.False;
                }
            }
            wc.Close();
            dgvAbitList.Update();
            dgvAbitList.ReadOnly = false;
            for (int i = 0; i < dgvAbitList.Rows.Count; i++)
            {
                dgvAbitList.Rows[i].ReadOnly = false;
                foreach (DataGridViewCell cell in dgvAbitList.Rows[i].Cells)
                    if (cell.Value is bool)
                    {
                        if (!(bool)cell.Value)
                            cell.ReadOnly = false;
                        else
                            cell.ReadOnly = true;
                    }
                    else
                        cell.ReadOnly = true;
            }
        }
        private string GetAbitFilterString()
        {
            string s = " WHERE 1=1 ";
            s += " AND ed.extEntry.StudyLevelGroupId IN (" + Util.BuildStringWithCollection(MainClass.lstStudyLevelGroupId) + ")";

            //обработали форму обучения  
            if (StudyFormId != null)
                s += " AND ed.extEntry.StudyFormId = " + StudyFormId;

            //обработали основу обучения  
            if (StudyBasisId != null)
                s += " AND ed.extEntry.StudyBasisId = " + StudyBasisId;

            if (rbCommon.Checked)
                s += " and ed.extentry.IsForeign = 0 and extentry.isCrimea= 0 ";
            else if (rbForeign.Checked)
                s += " and ed.extentry.IsForeign = 1 and extentry.isCrimea= 0 ";
            else if (rbCrimea.Checked)
                s += " and ed.extentry.IsForeign = 0 and extentry.isCrimea= 1 ";
            return s;
        }
        private void btnFillGrid_Click(object sender, EventArgs e)
        {
            FillGrid(true);
        }
        private void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillLicenseProgram();
            FillGrid(false);
        }
        private void cbStudyBasis_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillStudyForm();
        }
        private void cbStudyFor_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillLicenseProgram();
        }
        private void cbLicenseProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillGrid(false);
        }

        private bool UpdateResult()
        {
            bool result = false;
            foreach (DataGridViewRow dw in dgvAbitList.Rows)
            {
                foreach (DataGridViewCell dcell in dw.Cells)
                {
                    if (dcell.Value is bool)
                        continue;

                    if (dcell.Style.BackColor == Color.LightGreen)
                    {
                        string PersId = dcell.Value.ToString();
                        PersId = PersId.Substring(PersId.IndexOf("_") + 1);
                        PersId = PersId.Substring(0, PersId.LastIndexOf("_"));
                        PersId = PersId.Substring(PersId.IndexOf("_") + 1);
                    }
                }
            }
            return result;
        }
        private void DeletePaintGrid()
        {
            foreach (DataGridViewRow rw in dgvAbitList.Rows)
            {
                foreach (DataGridViewCell cl in rw.Cells)
                {
                    cl.Style.BackColor = Color.Empty;
                }
            }
        }
        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if (dgvAbitList.Rows.Count > 0)
            {
                DataTable tbl = ((DataView)dgvAbitList.DataSource).Table.Copy();

                string sheetName = "export";

                List<string> ColumnsNotVisibleList = new List<string>();
                List<int> ColumnsIndex = new List<int>();
                foreach (DataGridViewColumn col in dgvAbitList.Columns)
                    if (!col.Visible || (dgvAbitList.Rows[0].Cells[col.Name].Value is bool))
                    {
                        ColumnsNotVisibleList.Add(col.Name);
                    }
                    else
                    {
                        ColumnsIndex.Add(dgvAbitList.Columns.IndexOf(col));
                    }


                foreach (string s in ColumnsNotVisibleList)
                if (tbl.Columns.Contains(s))
                {
                    tbl.Columns.Remove(s);
                }


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
                        Excel.Range Range3 = ws.Range[ws.Cells[1, 1], ws.Cells[6, tbl.Columns.Count]];
                        Range3.WrapText = true;
                        Range3.RowHeight = rowHeight;
                        Range3.ColumnWidth = colFIOWidth;
                        Range3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        Range3.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                        Range3 = ws.Range[ws.Cells[7, 1], ws.Cells[tbl.Rows.Count, tbl.Columns.Count]];
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
                                Color clr = dgvAbitList.Rows[rowindex+4].Cells[ColumnsIndex[colindex]].Style.BackColor;
                                if (clr != Color.Empty)
                                    Range3.Interior.Color = clr;
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
       
        private void dgvAbitList_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex < 1)
                return;
            // есть ли профили
            if (e.RowIndex < 0)
            {
                /* // профилей больше нет
                if (e.Button == MouseButtons.Right)
                {
                    if (!String.IsNullOrEmpty(dgvAbitList.Rows[RowInnerEntryInEntryId].Cells[e.ColumnIndex].Value.ToString()))
                    {
                        string query = @"select Id from ed.ProfileInObrazProgramInEntry where ObrazProgramInEntryId ='" + dgvAbitList.Rows[RowInnerEntryInEntryId].Cells[e.ColumnIndex].Value.ToString() + "'";
                        DataTable tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
                        if (tbl.Rows.Count < 2)
                            return;

                        dgvAbitList.CurrentCell = dgvAbitList.Rows[e.RowIndex].Cells[e.ColumnIndex];
                        ContextMenu m = new ContextMenu();
                        m.MenuItems.Add(new MenuItem("Открыть распределение по профилям", new EventHandler(this.ContextMenuProfile_OnClick)));
                        Point pCell = dgvAbitList.GetCellDisplayRectangle(dgvAbitList.CurrentCell.ColumnIndex, dgvAbitList.CurrentCell.RowIndex, true).Location;
                        Point pGrid = dgvAbitList.Location;
                        new Point(pCell.X + pGrid.X, pCell.Y + pGrid.Y + dgvAbitList.CurrentRow.Height);

                        m.Show(dgvAbitList, new Point(pCell.X + pGrid.X, pCell.Y + dgvAbitList.CurrentRow.Height));
                    }
                }
                */
            }
            // абитуриенты
            else
            {
                 if (e.Button == MouseButtons.Right)
                {
                    dgvAbitList.CurrentCell = dgvAbitList.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    ContextMenu m = new ContextMenu();
                    m.MenuItems.Add(new MenuItem("Перейти к зеленой позиции", new EventHandler(this.ContextMenuToGreen_OnClick)));
                    m.MenuItems.Add(new MenuItem("Открыть карточку абитуриента", new EventHandler(this.ContextMenuOpenCard_OnClick)));
                    m.MenuItems.Add(new MenuItem("Перейти к следующему конкурсу", new EventHandler(this.ContextMenuNextApp_OnClick)));

                    Point pCell = dgvAbitList.GetCellDisplayRectangle(dgvAbitList.CurrentCell.ColumnIndex, dgvAbitList.CurrentCell.RowIndex, true).Location;
                    Point pGrid = dgvAbitList.Location;
                    new Point(pCell.X + pGrid.X, pCell.Y + pGrid.Y + dgvAbitList.CurrentRow.Height);

                    m.Show(dgvAbitList, new Point(pCell.X + pGrid.X, pCell.Y + dgvAbitList.CurrentRow.Height));
                }
            }
        }
        private void ContextMenuToGreen_OnClick(object sender, EventArgs e)
        {
            string FIO = dgvAbitList.CurrentCell.Value.ToString().Substring(0, dgvAbitList.CurrentCell.Value.ToString().IndexOf('(') - 1);
            
        }
        private void ContextMenuOpenCard_OnClick(object sender, EventArgs e)
        {
            string FIO = dgvAbitList.CurrentCell.Value.ToString().Substring(0, dgvAbitList.CurrentCell.Value.ToString().IndexOf('(') - 1);
            
        }
        private void ContextMenuProfile_OnClick(object sender, EventArgs e)
        {
            /*
            int columnindex = dgvAbitList.CurrentCell.ColumnIndex;
            string EntryId = dgvAbitList.Rows[RowEntryId].Cells[columnindex].Value.ToString();
            string InnerEntryInEntryId = dgvAbitList.Rows[RowInnerEntryInEntryId].Cells[columnindex].Value.ToString();
            if (String.IsNullOrEmpty(InnerEntryInEntryId))
                return;
            List<Guid> PersonNumList = new List<Guid>();
            List<string> PersonFIOList = new List<string>();

            string value = "";
            for (int i = startrow; i < dgvAbitList.Rows.Count; i++)
            {
                if (dgvAbitList.Rows[i].Cells[columnindex].Style.BackColor == Color.LightGreen)
                   // (dgvAbitList.Rows[i].Cells[columnindex].Style.BackColor == Color.LightBlue))
                {
                    value = dgvAbitList.Rows[i].Cells[columnindex].Value.ToString();
                    if (String.IsNullOrEmpty(value))
                    {
                        break;
                    }
                    string NUMFIO = value.Substring(0, value.IndexOf("(") - 1);
                    int index = PersonListFio.IndexOf(NUMFIO);
                    if (index > -1)
                    {
                        PersonNumList.Add(PersonList[index]);
                        PersonFIOList.Add(NUMFIO);
                    }
                    else
                        MessageBox.Show(this, "SomeError while searching FIO and Person.Id: " + value, "ContextMenuProfile_OnClick");
                }
            }
            new MyListRatingProfileList(ObrazProgramInEntryId, EntryId, PersonNumList, PersonFIOList, btnGreenIsClicked).Show();
             */
        }
        private void ContextMenuNextApp_OnClick(object sender, EventArgs e)
        {
            string FIO = dgvAbitList.CurrentCell.Value.ToString().Substring(0, dgvAbitList.CurrentCell.Value.ToString().IndexOf('(') - 1);
        }
        private void tbAbitsTop_MouseClick(object sender, MouseEventArgs e)
        {
            rbAbitsTop.Checked = true;
        }
       
        private void FillGrid(bool update)
        {
            if (update)
                FillGrid(GetAbitFilterString());

            if (FacultyId != null)
            {
                foreach (DataGridViewColumn clm in dgvAbitList.Columns)
                {
                    if (dgvAbitList.Rows[0].Cells[clm.Index].Value.ToString() != FacultyId.Value.ToString())
                    {
                        clm.Visible = false;
                    }
                    else
                    {
                        dgvAbitList.Columns[clm.Index - 1].Visible = true;
                        clm.Visible = true;
                    }
                }
            }
            else
            {
                foreach (DataGridViewColumn clm in dgvAbitList.Columns)
                {
                    clm.Visible = true;
                }
            }

            if (LicenseProgramId != null)
            {
                foreach (DataGridViewColumn clm in dgvAbitList.Columns)
                {
                    if (clm.Visible == true)
                    {
                        if (dgvAbitList.Rows[0].Cells[clm.Index].Value.ToString() != ((KeyValuePair<string, string>)cbLicenseProgram.SelectedItem).Value)
                        {
                            clm.Visible = false;
                        }
                        else
                        {
                            dgvAbitList.Columns[clm.Index - 1].Visible = true;
                            clm.Visible = true;
                        }
                    }
                }
            }
            if (dgvAbitList.Columns.Contains("Id"))
                dgvAbitList.Columns["Id"].Visible = false;
            if (update)
                btn_GreenList.Enabled = true;
        }
        private void btnRePaint_Click(object sender, EventArgs e)
        {
            PaintGrid();
        }
        private void PaintGrid()
        {
            double DinamicWave = 1.0 * int.Parse(tbDinamicWave.Text) / 100;
            foreach (var x in EntryList)
            {
                x.SetMaxCountGreen(DinamicWave);
                x.SetIsGreen();
            }
            bool HasChanges = true;
            int whileind = 0;
            while (HasChanges && whileind < Coord.GetLengh())
            {
                HasChanges = false;
                whileind++;
                for (int entryid = 0; entryid < EntryList.Count; entryid++)
                {
                    clEntry entry = EntryList[entryid];
                    for (int abitid = 0; abitid < entry.Abits.Count; abitid++)
                    {
                        Abiturient ab = entry.Abits[abitid];
                        if (ab.IsGreen())
                        {
                            var lst = Coord.GetCoordintesList(ab.PersonId);
                            Coordinates c = lst.Where(t => t.entryindex == entryid && t.abitlistindex == abitid).Select(t => t).First();

                            for (int i = 0; i < lst.Count; i++)
                            {
                                Coordinates x = lst[i];
                                if (!x.InCompetition || (c.entryindex == x.entryindex && x.abitlistindex == c.abitlistindex))
                                    continue;

                                Abiturient tmp_ab = EntryList[x.entryindex].Abits[x.abitlistindex];
                                if (tmp_ab.IsGreen())
                                {
                                    if (tmp_ab.Priority == ab.Priority)
                                    {
                                        ab.SetIsRed();
                                        x.InCompetition = false;

                                        tmp_ab.SetIsRed();
                                    }
                                    else if (tmp_ab.Priority > ab.Priority)
                                    {
                                        tmp_ab.SetIsYellow();
                                        x.InCompetition = false;
                                        EntryList[x.entryindex].SetNextGreen();
                                    }
                                    else if (tmp_ab.Priority < ab.Priority)
                                    {
                                        ab.SetIsYellow();
                                        c.InCompetition = false;
                                        entry.SetNextGreen();
                                    }
                                    HasChanges = true;
                                }
                            }
                        }
                    }
                }
            }
            foreach (clEntry ent in EntryList)
            {
                for (int i = 0; i < ent.Abits.Count; i++)
                {
                    if (!ent.Abits[i].HasColor())
                        foreach (Column col in ent.ColumnList)
                            col.SetEmptyColor(i);
                    else if (ent.Abits[i].IsGreen())
                        foreach (Column col in ent.ColumnList)
                            col.SetGreenColor(i);
                    else if (ent.Abits[i].IsYellow())
                        foreach (Column col in ent.ColumnList)
                            col.SetYellowColor(i);
                    else if (ent.Abits[i].IsBeige())
                        foreach (Column col in ent.ColumnList)
                            col.SetBeigeColor(i);
                    else if (ent.Abits[i].IsBlue())
                        foreach (Column col in ent.ColumnList)
                            col.SetBlueColor(i);
                    else if (ent.Abits[i].IsRed())
                        foreach (Column col in ent.ColumnList)
                            col.SetRedColor(i);
                }
            }
        }
        private void PaintDataGridView()
        {
            foreach (clEntry ent in EntryList)
            {
                foreach (Column col in ent.ColumnList)
                {
                    for (int abitid = 0; abitid < ent.Abits.Count; abitid++)
                    {
                        dgvAbitList.Rows[abitid + LastSystemRow].Cells[col.ColumnIndex - 1].Style.BackColor =
                             dgvAbitList.Rows[abitid + LastSystemRow].Cells[col.ColumnIndex].Style.BackColor = col.GetAbitColor(abitid);
                    }
                }
            }
        }
        private void btnRestoreOriginals_Click(object sender, EventArgs e)
        {
            
        }
        private void tbDinamicWave_TextChanged(object sender, EventArgs e)
        {
            tbDinamicWave.Text = tbDinamicWave.Text.Replace('.', ',');
        }
        private bool Check()
        {
            int proc;
            if (!int.TryParse(tbDinamicWave.Text, out proc))
            {
                MessageBox.Show("Неверный формат процента зачисляемых","!");
                return false;
            }
            return true;
        }
    }

    public class PersonCoordinates
    {
        List<KeyValuePair<Guid, List<Coordinates>>> PersonCoordList;
        public PersonCoordinates()
        {
            PersonCoordList = new List<KeyValuePair<Guid, List<Coordinates>>>();
        }
        public void Add(Guid PersonId)
        {
            if (PersonCoordList.Where(x=>x.Key == PersonId).Count() == 0)
                PersonCoordList.Add(new KeyValuePair<Guid,List<Coordinates>>(PersonId, new List<Coordinates>()));
        }
        public void Add(Guid PersonId, Coordinates cord)
        {
            if (PersonCoordList.Where(x => x.Key == PersonId).Count() == 0)
                PersonCoordList.Add(new KeyValuePair<Guid, List<Coordinates>>(PersonId, new List<Coordinates>() { cord }));
            else
            {
                PersonCoordList.Where(x => x.Key == PersonId).First().Value.Add(cord);
            }
        }
        public List<Coordinates> GetCoordintesList(Guid PersonId)
        {
            return PersonCoordList.Where(x => x.Key == PersonId).Select(x => x.Value).First();
        }
        public int GetLengh()
        {
            return PersonCoordList.Count;
        }

    }

    public class Abiturient
    {
        public Guid AbitId;
        public Guid PersonId;
        public string regNum_FIO;
        public int Priority;
        private bool isGreen;
        private bool isBeige;
        private bool isYellow;
        private bool isBlue;
        private bool isRed;
        public bool HasOriginals;
        public Abiturient()
        {}
        public void SetIsBeige()
        {
            isGreen = isYellow = isBlue = false;
            isBeige = true;
        }
        public void SetIsGreen()
        {
            isBeige = isYellow = isBlue = false;
            isGreen = true;
        }
        public void SetIsYellow()
        {
            isBeige = isGreen = isBlue = false;
            isYellow = true;
        }
        public void SetIsBlue()
        {
            isBeige = isGreen = isYellow = false;
            isBlue = true;
        }
        public void SetIsRed()
        {
            isRed = true;
            isGreen = isBeige = isBlue = isYellow = false;
        }
        public bool IsGreen()
        {
            return isGreen;
        }
        public bool IsYellow()
        {
            return isYellow;
        }
        public bool IsBeige()
        {
            return isBeige;
        }
        public bool IsBlue()
        {
            return isBlue;
        }
        public bool IsRed()
        {
            return isRed;
        }

        public bool HasColor()
        {
            return isGreen || isBlue || isBeige || isYellow || isRed;
        }
    }
    public class Column
    {
        public List<int> InnerPriorities;
         List<Color> AbitsColorList;
        public string InnerEntryId;
        public string InnerObrazProgramName;
        public string ProfileName;
        public int MaxKCP;
        public string ColumnName;
        public int ColumnIndex;

        public Column()
        {
            ColumnName = Guid.NewGuid().ToString();
            InnerPriorities = new List<int>();
            AbitsColorList = new List<Color>();
        }
        public void AbitColorListAdd()
        {
            AbitsColorList.Add(Color.Empty);
        }
        public Color GetAbitColor(int abitid)
        {
            return AbitsColorList[abitid];
        }
        public void SetEmptyColor(int abitid)
        {
            AbitsColorList[abitid] = Color.Empty;
        }
        public void SetGreenColor(int abitid)
        {
            AbitsColorList[abitid] = Color.LightGreen;
        }
        public void SetYellowColor(int abitid)
        {
            AbitsColorList[abitid] = Color.Yellow;
        }
        public void SetBlueColor(int abitid)
        {
            AbitsColorList[abitid] = Color.LightBlue ;
        }
        public void SetRedColor(int abitid)
        {
            AbitsColorList[abitid] = Color.Red;
        }
        public void SetBeigeColor(int abitid)
        {
            AbitsColorList[abitid] = Color.Beige;
        } 
    }
    public class clEntry
    {
        public List<Column> ColumnList;
        public List<Abiturient> Abits;
        public string EntryId;
        public int FacultyId;
        public int KCP;
        public int MaxCountGreen;
        public string LicenseProgramName;
        public string ObrazProgramName;
        public string StudyBasis;
        public string StudyForm;

        public clEntry()
        {
            ColumnList = new List<Column>();
            Abits = new List<Abiturient>();
        }

        public clEntry(string licprog, string obrazprog, string prof)
        {
            ColumnList = new List<Column>();
            Abits = new List<Abiturient>();
            LicenseProgramName = licprog;
            ObrazProgramName = obrazprog;
        }
        public void AddColumn(Column cl)
        {
            ColumnList.Add(cl);
        }
        public void SetIsGreen()
        {
            for (int i = 0; (i < MaxCountGreen)&&(i<Abits.Count); i++)
                Abits[i].SetIsBeige();
            int j = 0;
            int a = 0;
            while (j < MaxCountGreen && a < Abits.Count)
            {
                if (Abits[a].HasOriginals)
                {
                    Abits[a].SetIsGreen();
                    j++;
                }
                a++;
            }
        }
        public void SetMaxCountGreen(double Persent)
        {
            MaxCountGreen = (int)Math.Ceiling(Persent * KCP);
        }
        public void SetNextGreen()
        {
            foreach (Abiturient x in Abits)
            {
                if (!x.HasColor() && x.HasOriginals)
                {
                    x.SetIsGreen();
                    return;
                }
            }
        }
    }
    public struct Coordinates
    {
        public int entryindex;
        public int abitlistindex;
        public bool InCompetition;
    }
}