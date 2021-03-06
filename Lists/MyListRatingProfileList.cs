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
using PriemLib;

namespace Priem
{
    public partial class MyListRatingProfileList : BookList
    {
        Guid _ObrazProgramInEntryId;
        Guid _EntryId;
        List<Guid> PersonNumList = new List<Guid>();
        List<string> PersonFIOList = new List<string>();
        int startrow = 4;
        bool IsGreen;

        public MyListRatingProfileList(string Id, string EntryId, List<Guid> List, List<string> ListFio, bool isgr)
        {
            InitializeComponent();
            Dgv = dgvAbitProfileList;

            _ObrazProgramInEntryId = Guid.Parse(Id);
            _EntryId = Guid.Parse(EntryId);
            PersonNumList = List;
            PersonFIOList = ListFio;
            IsGreen = isgr;
            InitControls();
        }

        public override void UpdateDataGrid()
        {
        }

        protected override void ExtraInit()
        {
            base.ExtraInit();
            btnRemove.Visible = btnAdd.Visible = false;
            btnExcel.Enabled = false;
            btnGreenList.Visible = false;
            btnGreenList.Enabled = IsGreen;

            if (MainClass.IsOwner())
                btnGreenList.Visible = true;

            if (MainClass.dbType == PriemType.PriemMag)
            {
                pictureBoxYellow.Location = pictureBoxLightGreen.Location;
                pictureBoxLightGreen.Location = pictureBoxLightBlue.Location;
                pictureBoxLightBlue.Location = pictureBoxThistle.Location;
                pictureBoxThistle.Visible = false;
                labelYellow.Location = labelLightGreen.Location;
                labelLightGreen.Location = labelLightBlue.Location;
                labelLightBlue.Location = labelThistle.Location;
                labelThistle.Visible = false;
            }
            
            _title = "Рейтинговый список с внутренними приоритетами";
            try
            {
                string query = @" select qEntry.FacultyName, qEntry.LicenseProgramName, qEntry.StudyBasisName, qEntry.StudyFormName, SP_ObrazProgram.Name  
                                  from ed.ObrazProgramInEntry 
                                  inner join ed.qEntry on qEntry.Id = EntryId
                                  inner join ed.SP_ObrazProgram on ObrazProgramInEntry.ObrazProgramId= SP_ObrazProgram.Id
                                  where ObrazProgramInEntry.Id = '" + _ObrazProgramInEntryId + "'";
                DataTable tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
                if (tbl.Rows.Count == 1)
                {
                    DataRow rw = tbl.Rows[0];
                    tbFaculty.Text = rw.Field<string>("FacultyName");
                    tbLicenseProgram.Text = rw.Field<string>("LicenseProgramName");
                    tbStudyBasis.Text = rw.Field<string>("StudyBasisName");
                    tbObrazProgramInEntry.Text = rw.Field<string>("Name");
                    tbStudyForm.Text = rw.Field<string>("StudyFormName");
                } 
                FillGrid();
            }
            catch (Exception exc)
            {
                WinFormsServ.Error("Ошибка при инициализации формы " + exc.Message);
            }
        }
        private void FillGrid()
        {
            string query = @"select distinct ProfileInObrazProgramInEntry.Id as Id, SP_Profile.Name as Name, ProfileInObrazProgramInEntry.KCP, ProfileInObrazProgramInEntry.EgeExamNameId ,EgeExamName.Name as EgeName, ProfileInObrazProgramInEntry.EgeMin
                                from ed.ProfileInObrazProgramInEntry 
                                inner join ed.SP_Profile on SP_Profile.Id = ProfileInObrazProgramInEntry.ProfileId 
                                left join ed.EgeExamName on EgeExamName.Id = EgeExamNameId
                                where 
                                ObrazProgramInEntryId ='" + _ObrazProgramInEntryId + @"' 
                                ";
            DataTable tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            if (tbl.Rows.Count == 0)
                return;

            DataTable examTable = new DataTable();
            examTable.Columns.Add("Id");

            DataColumn clm;
            DataRow rowProfileName = examTable.NewRow();
            DataRow rowProfileId = examTable.NewRow();
            DataRow rowKCP = examTable.NewRow();
            DataRow rowEge = examTable.NewRow();

            foreach (DataRow rw_profile in tbl.Rows)
            {
                clm = new DataColumn();
                String ColName = rw_profile.Field<Guid>("Id").ToString();
                clm.ColumnName = ColName;
                examTable.Columns.Add(clm);
                rowProfileName[ColName] = rw_profile.Field<string>("Name");
                rowProfileId[ColName] = ColName;
                rowKCP[ColName] = rw_profile.Field<int>("KCP");
                rowEge[ColName] = (!String.IsNullOrEmpty(rw_profile.Field<int?>("EgeExamNameId").ToString())) ? rw_profile.Field<int?>("EgeExamNameId") + "_" + rw_profile.Field<string>("EgeName") + "(" + rw_profile.Field<int?>("EgeMin") + ")" : "";
            }


            for (int i = 1; i < examTable.Columns.Count; i++)
            {
                int kcp_new = int.Parse(rowKCP[i].ToString());
                rowKCP[i] = kcp_new - int.Parse(MainClass.Bdc.GetStringValue(@"
                                select COUNT(extEntryView.Id) 
                                from ed.extEntryView
                                inner join ed.Abiturient on AbiturientId = Abiturient.Id
                                where Abiturient.EntryId = '" + _EntryId + @"' and 
                                Abiturient.ObrazProgramInEntryId = '" + _ObrazProgramInEntryId + @"' and 
                                Abiturient.ProfileInObrazProgramInEntryId = '" + rowProfileId[i].ToString() + @"'
                                and Abiturient.CompetitionId NOT IN (11,12)"));
            } 

            examTable.Rows.Add(rowProfileName);
            examTable.Rows.Add(rowProfileId);
            examTable.Rows.Add(rowEge);
            examTable.Rows.Add(rowKCP);

            // abiturients 
            query = @"select distinct
                          ApplicationDetails.ProfileInObrazProgramInEntryId
                         ,ApplicationDetails.ProfileInObrazProgramInEntryPriority
                         from ed.Abiturient  
                         left join ed.ApplicationDetails on ApplicationDetails.ApplicationId = Abiturient.Id
  
                         where Abiturient.PersonId=@PersonId and Abiturient.BackDoc = 0 and Abiturient.IsGosLine=0 
                         and ApplicationDetails.ObrazProgramInEntryId='" + _ObrazProgramInEntryId + "'";
            DataRow rw_list;
            foreach (Guid PersonNum in PersonNumList)
            {
                rw_list = examTable.NewRow();
                foreach (DataColumn column in examTable.Columns)
                {
                    rw_list[column.ColumnName] = PersonNum.ToString();
                }
                DataSet ds = MainClass.Bdc.GetDataSet(query, new SortedList<string, object> { { "@PersonId", PersonNum } });
                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    String ColName = row.Field<Guid>("ProfileInObrazProgramInEntryId").ToString();
                    int Priority = row.Field<int>("ProfileInObrazProgramInEntryPriority");
                    rw_list[ColName] += "_" + Priority.ToString();
                }
                examTable.Rows.Add(rw_list);
            }

            DataView dv = new DataView(examTable);

            dgvAbitProfileList.DataSource = dv;
            dgvAbitProfileList.AllowUserToOrderColumns = false;
            for (int i = 0; i < dgvAbitProfileList.Columns.Count; i++)
                dgvAbitProfileList.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            dgvAbitProfileList.ColumnHeadersVisible = false;
            dgvAbitProfileList.Columns["Id"].Visible = false;
            dgvAbitProfileList.Update();
            //GridPaint();
            //CopyTable();
        }
        private void GridPaint()
        {
            int startcol = 1;
            // сначала все КЦП зеленые
            for (int colindex = startcol; colindex < dgvAbitProfileList.Columns.Count; colindex++)
            {
                int KCP = 0;
                int.TryParse(dgvAbitProfileList.Rows[startrow - 1].Cells[colindex].Value.ToString(), out KCP);

                for (int j = startrow; (j < KCP + startrow) && (j < dgvAbitProfileList.Rows.Count); j++)
                {
                    if (String.IsNullOrEmpty(dgvAbitProfileList.Rows[j].Cells[colindex].Value.ToString()))
                        break;
                    dgvAbitProfileList.Rows[j].Cells[colindex].Style.BackColor = Color.LightGreen;
                }
            }
            // лица без приоритетов голубенькие
            for (int colindex = startcol; colindex < dgvAbitProfileList.Columns.Count; colindex++)
            {
                for (int j = startrow; j < dgvAbitProfileList.Rows.Count; j++)
                {
                    if (dgvAbitProfileList.Rows[j].Cells[colindex].Value.ToString().EndsWith("_0"))
                    {
                        dgvAbitProfileList.Rows[j].Cells[colindex].Style.BackColor = Color.LightBlue;
                    }
                }
            }
            // если для профиля есть конкретный язык, нужно сделать отсев
            for (int colindex = startcol; colindex < dgvAbitProfileList.Columns.Count; colindex++)
            {
                if (String.IsNullOrEmpty(dgvAbitProfileList.Rows[startrow - 2].Cells[colindex].Value.ToString()))
                    continue;

                string EgeExamNameId = dgvAbitProfileList.Rows[startrow - 2].Cells[colindex].Value.ToString();
                EgeExamNameId = EgeExamNameId.Substring(0, EgeExamNameId.IndexOf("_"));
                string sEgeMin = dgvAbitProfileList.Rows[startrow - 2].Cells[colindex].Value.ToString();
                sEgeMin = sEgeMin.Substring(sEgeMin.IndexOf("(") + 1);
                sEgeMin = sEgeMin.Substring(0, sEgeMin.IndexOf(")"));
                int EgeMin = int.Parse(sEgeMin);

                int KCP_temp = 0;
                if (int.TryParse(dgvAbitProfileList.Rows[startrow - 1].Cells[colindex].Value.ToString(), out KCP_temp))

                    for (int j = startrow; j < dgvAbitProfileList.Rows.Count; j++)
                    {
                        DataGridViewCell cell = dgvAbitProfileList.Rows[j].Cells[colindex];
                        string cellvalue = cell.Value.ToString().Substring(0, cell.Value.ToString().IndexOf("_"));

                        int EgeAbitValue = (int?)MainClass.Bdc.GetValue("select Value from ed.extEgeMark where PersonId = '" + cellvalue + "' and EgeExamNameId=" + EgeExamNameId + " and FBSStatusId=1") ?? 0;
                        if (EgeAbitValue < EgeMin)
                        {
                            if ((dgvAbitProfileList.Rows[j].Cells[colindex].Style.BackColor == Color.LightGreen) || (dgvAbitProfileList.Rows[j].Cells[colindex].Style.BackColor == Color.LightBlue))
                            {
                                // сдвинуть зеленку;
                                for (int row_temp = startrow + KCP_temp; row_temp < dgvAbitProfileList.Rows.Count; row_temp++)
                                {
                                    if (String.IsNullOrEmpty(dgvAbitProfileList.Rows[row_temp].Cells[colindex].Value.ToString()))
                                        break;
                                    if (dgvAbitProfileList.Rows[row_temp].Cells[colindex].Style.BackColor == Color.Empty)
                                    {
                                        dgvAbitProfileList.Rows[row_temp].Cells[colindex].Style.BackColor = Color.LightGreen;
                                        break;
                                    }
                                }
                            }
                            dgvAbitProfileList.Rows[j].Cells[colindex].Style.BackColor = Color.Thistle;
                        }
                    }
            }

            // по готовенькому расставить приоритетики
            for (int rowindex = startrow; rowindex < dgvAbitProfileList.Rows.Count; rowindex++)
            {
                DataGridViewRow row = dgvAbitProfileList.Rows[rowindex];
                if (String.IsNullOrEmpty(row.Cells[startcol].Value.ToString()))
                    break;
                for (int i = startcol; i < dgvAbitProfileList.Columns.Count; i++)
                {
                    DataGridViewCell cell = row.Cells[i];

                    if (cell.Style.BackColor != Color.LightGreen)
                        continue;

                    int priority = 0;
                    if (!int.TryParse(cell.Value.ToString().Substring(cell.Value.ToString().IndexOf("_") + 1), out priority))
                    {
                        MessageBox.Show("Ошибка при проверке приоритета (перерисовка)");
                    }

                    for (int j = i + 1; j < dgvAbitProfileList.Columns.Count; j++)
                    {
                        DataGridViewCell temp_cell = row.Cells[j];
                        if (temp_cell.Style.BackColor == Color.LightBlue)
                            continue;

                        int temp_priority = 0;
                        if (!int.TryParse(temp_cell.Value.ToString().Substring(temp_cell.Value.ToString().IndexOf("_") + 1), out temp_priority))
                        {
                            MessageBox.Show("Ошибка при проверке приоритета (перерисовка)");
                        }
                        if (temp_priority > priority)
                        {
                            if (temp_cell.Style.BackColor == Color.LightGreen)
                            {
                                // сдвинуть зеленку;
                                int KCP_temp = 0;
                                if (int.TryParse(dgvAbitProfileList.Rows[startrow - 1].Cells[j].Value.ToString(), out KCP_temp))
                                {
                                    for (int row_temp = startrow + KCP_temp; row_temp < dgvAbitProfileList.Rows.Count; row_temp++)
                                    {
                                        if (String.IsNullOrEmpty(dgvAbitProfileList.Rows[row_temp].Cells[j].Value.ToString()))
                                            break;
                                        if (dgvAbitProfileList.Rows[row_temp].Cells[j].Style.BackColor == Color.Empty)
                                        {
                                            dgvAbitProfileList.Rows[row_temp].Cells[j].Style.BackColor = Color.LightGreen;
                                            break;
                                        }
                                    }
                                }
                            }
                            if ((temp_cell.Style.BackColor == Color.LightGreen) || (temp_cell.Style.BackColor == Color.LightBlue) || (temp_cell.Style.BackColor == Color.Empty))
                                temp_cell.Style.BackColor = Color.Yellow;
                        }
                        else
                        {
                            if (priority == temp_priority)
                            {
                                string cellvalue = cell.Value.ToString().Substring(0, cell.Value.ToString().IndexOf("_"));
                                string FIO = PersonFIOList[PersonNumList.IndexOf(Guid.Parse(cellvalue))];
                                MessageBox.Show(this, "Вы знаете, у абитуриента: " + FIO + " существуют повторяющиеся приоритеты", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            if (temp_cell.Style.BackColor == Color.LightGreen)
                            {
                                priority = temp_priority;
                                cell.Style.BackColor = Color.Yellow;
                                // спустить зеленку
                                int KCP_temp = 0;
                                if (int.TryParse(dgvAbitProfileList.Rows[startrow - 1].Cells[cell.ColumnIndex].Value.ToString(), out KCP_temp))
                                {
                                    for (int row_temp = startrow + KCP_temp; row_temp < dgvAbitProfileList.Rows.Count; row_temp++)
                                    {
                                        if (String.IsNullOrEmpty(dgvAbitProfileList.Rows[row_temp].Cells[cell.ColumnIndex].Value.ToString()))
                                            break;
                                        if (dgvAbitProfileList.Rows[row_temp].Cells[cell.ColumnIndex].Style.BackColor == Color.Empty)
                                        {
                                            dgvAbitProfileList.Rows[row_temp].Cells[cell.ColumnIndex].Style.BackColor = Color.LightGreen;
                                            break;
                                        }
                                    }
                                }
                                cell = row.Cells[j];
                            }
                        }
                    } 
                    foreach (DataGridViewCell cells in row.Cells)
                    {
                        if (cell.ColumnIndex == startcol - 1)
                            continue;

                        if (cells != cell)
                        {
                            if (cells.Style.BackColor == Color.Empty)
                            { 
                                cells.Style.BackColor = Color.Yellow;
                            }
                        }
                    } 
                    break;
                }
            }
        }

        private void CopyTable()
        {
            int startcol = 1;

            for (int j = startcol; j < dgvAbitProfileList.Columns.Count; j++)
            {
                string value = dgvAbitProfileList.Rows[startrow - 2].Cells[j].Value.ToString();
                if (!String.IsNullOrEmpty(value))
                    dgvAbitProfileList.Rows[startrow - 2].Cells[j].Value = value.Substring(value.IndexOf("_") + 1);
            }
            for (int i = startrow; i < dgvAbitProfileList.Rows.Count; i++)
            {
                for (int j = startcol; j < dgvAbitProfileList.Columns.Count; j++)
                {
                    string value = dgvAbitProfileList.Rows[i].Cells[j].Value.ToString();
                    Guid PersId = Guid.Parse(value.Substring(0, value.IndexOf("_")));
                    dgvAbitProfileList.Rows[i].Cells[j].Value = PersonFIOList[PersonNumList.IndexOf(PersId)] + " (" + value.Substring(value.IndexOf("_") + 1) + ")";
                }
            }
            // ProfileId
            dgvAbitProfileList.Rows[1].Visible = false;
        }

        private void dgvAbitProfileList_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex < 1)
                return;
            // есть ли профили
            if (e.RowIndex < 2)
                return;
            // абитуриенты
            else
                if (e.Button == MouseButtons.Right)
                {
                    dgvAbitProfileList.CurrentCell = dgvAbitProfileList.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    ContextMenu m = new ContextMenu();
                    m.MenuItems.Add(new MenuItem("Перейти к зеленой позиции", new EventHandler(this.ContextMenuToGreen_OnClick)));
                    m.MenuItems.Add(new MenuItem("Открыть карточку абитуриента", new EventHandler(this.ContextMenuOpenCard_OnClick)));

                    Point pCell = dgvAbitProfileList.GetCellDisplayRectangle(dgvAbitProfileList.CurrentCell.ColumnIndex, dgvAbitProfileList.CurrentCell.RowIndex, true).Location;
                    Point pGrid = dgvAbitProfileList.Location;
                    new Point(pCell.X + pGrid.X, pCell.Y + pGrid.Y + dgvAbitProfileList.CurrentRow.Height);

                    m.Show(dgvAbitProfileList, new Point(pCell.X + pGrid.X, pCell.Y + dgvAbitProfileList.CurrentRow.Height));
                }
        }
        private void ContextMenuToGreen_OnClick(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in dgvAbitProfileList.CurrentRow.Cells)
            {
                if (cell.Style.BackColor == Color.LightGreen)
                {
                    dgvAbitProfileList.CurrentCell = cell;
                }
            }
        }
        private void ContextMenuOpenCard_OnClick(object sender, EventArgs e)
        {
            string FIO = dgvAbitProfileList.CurrentCell.Value.ToString().Substring(0, dgvAbitProfileList.CurrentCell.Value.ToString().IndexOf('(') - 1);
            int index = PersonFIOList.IndexOf(FIO);

            MainClass.OpenCardPerson(PersonNumList[index].ToString(), this, dgvAbitProfileList.CurrentRow.Index);
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            if (dgvAbitProfileList.Rows.Count > 2)
            {
                DataTable tbl = ((DataView)dgvAbitProfileList.DataSource).Table.Copy();

                string sheetName = "export";
                if (tbl.Columns.Contains("Id"))
                {
                    tbl.Columns.Remove("Id");
                }

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

                        Excel.Range Range3 = ws.Range[ws.Cells[1, 1], ws.Cells[2, tbl.Columns.Count]];
                        Range3.WrapText = true;
                        Range3.RowHeight = rowHeight;
                        Range3.ColumnWidth = colFIOWidth;
                        Range3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        Range3.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                        Range3 = ws.Range[ws.Cells[3, 1], ws.Cells[tbl.Rows.Count, tbl.Columns.Count]];
                        Range3.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                        for (int rowindex = 0; rowindex < tbl.Rows.Count; rowindex++)
                        {
                            DataRow dr = tbl.Rows[rowindex];
                            j = 1;
                            for (int colindex = 0; colindex < tbl.Columns.Count; colindex++)
                            {
                                DataColumn dc = tbl.Columns[colindex];
                                ws.Cells[i, j] = dr[dc.ColumnName] == null ? "" : dr[dc.ColumnName].ToString();
                                Range3 = ws.Cells[i, j];
                                Color clr = dgvAbitProfileList.Rows[rowindex].Cells[colindex + 1].Style.BackColor;
                                if (clr != Color.Empty)
                                    Range3.Interior.Color = dgvAbitProfileList.Rows[rowindex].Cells[colindex + 1].Style.BackColor;
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

        private void btnGreenList_Click(object sender, EventArgs e)
        {
            int startcol = 1;
            NewWatch wc = new NewWatch();
            wc.Show();
            wc.SetText("Обновление данных...");
            wc.SetMax(dgvAbitProfileList.Columns.Count);
            for (int clmn = startcol; clmn < dgvAbitProfileList.Columns.Count; clmn++)
            {
                // 0 ProfileName
                // 1 ProfileInObrazProgramInEntryId
                // 2 KCP 
                // 3 абитуриентик
                string ObrazProgramInEntryId = _ObrazProgramInEntryId.ToString();
                string ProfileId = dgvAbitProfileList.Rows[1].Cells[clmn].Value.ToString();
                string PersonId = "";
                string AbitId = "";
                string NumFio = "";
                string query = @"select Abiturient.Id from ed.Abiturient where EntryId ='" + _EntryId.ToString() + "' and PersonId='";


                for (int rowindex = startrow; rowindex < dgvAbitProfileList.Rows.Count; rowindex++)
                {
                    string value = dgvAbitProfileList.Rows[rowindex].Cells[clmn].Value.ToString();
                    if (String.IsNullOrEmpty(value))
                        break;

                    if (dgvAbitProfileList.Rows[rowindex].Cells[clmn].Style.BackColor == Color.Empty)
                        break;

                    if ((dgvAbitProfileList.Rows[rowindex].Cells[clmn].Style.BackColor == Color.LightGreen) ||
                        (dgvAbitProfileList.Rows[rowindex].Cells[clmn].Style.BackColor == Color.LightBlue))
                    {
                        NumFio = value.Substring(0, value.IndexOf("(") - 1);
                        int personIndexInList = PersonFIOList.IndexOf(NumFio);
                        if (personIndexInList > -1)
                        {
                            PersonId = PersonNumList[personIndexInList].ToString();
                            AbitId = MainClass.Bdc.GetDataSet(query + PersonId + "'").Tables[0].Rows[0].Field<Guid>("Id").ToString();
                            if (!String.IsNullOrEmpty(AbitId))
                            {
                                MainClass.Bdc.ExecuteQuery("Update ed.AbiturientGreen  set ProfileInObrazProgramInEntryId = '" + ProfileId + "' where AbiturientId ='" + AbitId + "' and ObrazProgramInEntryId ='" + ObrazProgramInEntryId + "'");
                            }
                            else
                            {
                                MessageBox.Show("Ошибка в процессе получения AbiturientId (btn_GreenList_Click)");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Ошибка в процессе получения PersonId (btn_GreenList_Click)");
                        }
                    }
                }

                wc.PerformStep();
                wc.SetText("Обновление данных: Обработано конкурсов " + clmn + "/" + (dgvAbitProfileList.Columns.Count - 1) + "...");
            }
            wc.Close();
        }

        private void MyListRatingProfileList_Shown(object sender, EventArgs e)
        {
            GridPaint();
            CopyTable();
            btnExcel.Enabled = true;
        }

        protected override void OpenCard(string itemId)
        {
            //base.OpenCard(itemId);
        }

        private void dgvAbitProfileList_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < startrow)
                return;
            if (e.ColumnIndex < 1)
                return;

            string FIO = dgvAbitProfileList.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Substring(0, dgvAbitProfileList.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().IndexOf('(') - 1);
            int index = PersonFIOList.IndexOf(FIO);

            MainClass.OpenCardPerson(PersonNumList[index].ToString(), this, dgvAbitProfileList.CurrentRow.Index);
        }
    }
}