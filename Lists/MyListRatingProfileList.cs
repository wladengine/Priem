﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using WordOut;
using EducServLib;
using BDClassLib;
using BaseFormsLib;

namespace Priem 
{
    public partial class MyListRatingProfileList : BookList
    {
        Guid _ObrazProgramInEntryId;
        List<Guid> PersonNumList = new List<Guid>();
        List<string> PersonFIOList = new List<string>();


        public MyListRatingProfileList(string Id, List<Guid>List, List<string>ListFio)
        {
            InitializeComponent();
            Dgv = dgvAbitProfileList;
            _ObrazProgramInEntryId = Guid.Parse(Id);
            PersonNumList = List;
            PersonFIOList = ListFio;
            InitControls();
            
        }

        public override void UpdateDataGrid()
        {
        }

        protected override void ExtraInit()
        {
            base.ExtraInit();
            btnRemove.Visible = btnAdd.Visible = false;
            _title = "Рейтинговый список с внутренними приоритетами";
            try
            {
                string query = @" select FacultyName, LicenseProgramName, ObrazProgramName, StudyBasisName, StudyFormName  
                                  from ed.ObrazProgramInEntry 
                                  inner join ed.qEntry on qEntry.Id = EntryId
                                  where ObrazProgramInEntry.Id = '" + _ObrazProgramInEntryId + "'";
                DataTable tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
                if (tbl.Rows.Count == 1)
                {
                    DataRow rw = tbl.Rows[0];
                    tbFaculty.Text = rw.Field<string>("FacultyName");
                    tbLicenseProgram.Text = rw.Field<string>("LicenseProgramName");
                    tbObrazProgram.Text = rw.Field<string>("ObrazProgramName");
                    tbStudyBasis.Text = rw.Field<string>("StudyBasisName");
                    tbStudyForm.Text = rw.Field<string>("StudyFormName");
                }
                query = @"select SP_ObrazProgram.Name
                          from   ed.ObrazProgramInEntry
                          inner join ed.SP_ObrazProgram on ObrazProgramId= SP_ObrazProgram.Id
                          where ObrazProgramInEntry.Id = '" + _ObrazProgramInEntryId + "'";
                tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
                if (tbl.Rows.Count == 1)
                {
                    DataRow rw = tbl.Rows[0];
                    tbObrazProgramInEntry.Text = rw.Field<string>("Name");
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
            string query = @"select distinct ProfileInObrazProgramInEntry.Id as Id, SP_Profile.Name as Name, ProfileInObrazProgramInEntry.KCP
                                from ed.ProfileInObrazProgramInEntry 
                                inner join ed.SP_Profile on SP_Profile.Id = ProfileInObrazProgramInEntry.ProfileId 
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
            DataRow rowKCP = examTable.NewRow();

            foreach (DataRow rw_profile in tbl.Rows)
            {
                clm = new DataColumn();
                String ColName = rw_profile.Field<Guid>("Id").ToString();
                clm.ColumnName = ColName;
                examTable.Columns.Add(clm);
                rowProfileName[ColName] = rw_profile.Field<string>("Name");
                rowKCP[ColName] = rw_profile.Field<int>("KCP");
            }
            examTable.Rows.Add(rowProfileName);
            examTable.Rows.Add(rowKCP);
            // abiturients 
            query = @"select 
                          ApplicationDetails.ProfileInObrazProgramInEntryId
                         ,ApplicationDetails.ProfileInObrazProgramInEntryPriority
                         from ed.Abiturient  
                         left join ed.ApplicationDetails on ApplicationDetails.ApplicationId = Abiturient.Id
  
                         where Abiturient.PersonId=@PersonId and Abiturient.BackDoc = 0 and Abiturient.IsGosLine=0 
                         and ApplicationDetails.ObrazProgramInEntryId='" + _ObrazProgramInEntryId+"'";
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
        private void GridPaint () 
        {
            int startrow = 2;
            int startcol = 1;
            for (int colindex = startcol; colindex < dgvAbitProfileList.Columns.Count; colindex++)
            {
                int KCP = 0;
                int.TryParse(dgvAbitProfileList.Rows[1].Cells[colindex].Value.ToString(), out KCP);

                for (int j = startrow; (j < KCP + startrow) && (j < dgvAbitProfileList.Rows.Count); j++)
                {
                    dgvAbitProfileList.Rows[j].Cells[colindex].Style.BackColor = Color.LightGreen;
                }
            }

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

            for (int rowindex = startrow; rowindex < dgvAbitProfileList.Rows.Count; rowindex++)
            {
                DataGridViewRow row = dgvAbitProfileList.Rows[rowindex];
                for (int i = startcol; i < dgvAbitProfileList.Columns.Count; i++)
                {
                    DataGridViewCell cell = row.Cells[i];
                    if (cell.Style.BackColor != Color.LightGreen)
                        continue; 

                    int priority = 0; 
                    if (!int.TryParse(cell.Value.ToString().Substring(cell.Value.ToString().IndexOf("_")+1), out priority))
                    {
                        MessageBox.Show("Ошибка при проверке приоритета (перерисовка)");
                    }

                    for (int j = i + 1; j < dgvAbitProfileList.Columns.Count; j++)
                    {
                        DataGridViewCell temp_cell = row.Cells[j];
                        if (temp_cell.Style.BackColor == Color.LightBlue)
                            continue;

                        int temp_priority = 0;
                        if (!int.TryParse(temp_cell.Value.ToString().Substring(temp_cell.Value.ToString().IndexOf("_")+1), out temp_priority))
                        {
                            MessageBox.Show("Ошибка при проверке приоритета (перерисовка)");
                        }
                        if (temp_priority > priority)
                        {
                            if (temp_cell.Style.BackColor == Color.LightGreen)
                            {
                                // сдвинуть зеленку;
                                int KCP_temp = 0;
                                if (int.TryParse(dgvAbitProfileList.Rows[1].Cells[j].Value.ToString(), out KCP_temp))
                                {
                                    for (int row_temp = startrow + KCP_temp; row_temp < dgvAbitProfileList.Rows.Count; row_temp++)
                                    {
                                        if ((dgvAbitProfileList.Rows[row_temp].Cells[j].Style.BackColor != Color.LightGreen) &&
                                       (dgvAbitProfileList.Rows[row_temp].Cells[j].Style.BackColor != Color.Yellow) &&
                                           (dgvAbitProfileList.Rows[row_temp].Cells[j].Style.BackColor != Color.LightBlue))
                                        //if (dgvAbitList.Rows[row_temp].Cells[kvp.Key].Style.BackColor == Color.White)
                                        {
                                            dgvAbitProfileList.Rows[row_temp].Cells[j].Style.BackColor = Color.LightGreen;
                                            break;
                                        }
                                    }
                                }
                            }
                            temp_cell.Style.BackColor = Color.Yellow;
                        }
                        else
                        { 
                            if (temp_cell.Style.BackColor == Color.LightGreen)
                            {
                                priority = temp_priority;
                                cell.Style.BackColor = Color.Yellow;
                                // спустить зеленку
                                int KCP_temp = 0;
                                if (int.TryParse(dgvAbitProfileList.Rows[1].Cells[cell.ColumnIndex].Value.ToString(), out KCP_temp))
                                {
                                    for (int row_temp = startrow + KCP_temp; row_temp < dgvAbitProfileList.Rows.Count; row_temp++)
                                    {
                                        if ((dgvAbitProfileList.Rows[row_temp].Cells[cell.ColumnIndex].Style.BackColor != Color.LightGreen) &&
                                       (dgvAbitProfileList.Rows[row_temp].Cells[cell.ColumnIndex].Style.BackColor != Color.Yellow) &&
                                           (dgvAbitProfileList.Rows[row_temp].Cells[cell.ColumnIndex].Style.BackColor != Color.LightBlue))
                                        //if (dgvAbitProfileList.Rows[row_temp].Cells[cell.ColumnIndex].Style.BackColor == Color.White)
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
                            if (cells.Style.BackColor != Color.LightBlue)
                            {
                                if (cells.Style.BackColor == Color.LightGreen)
                                {
                                    //  спустить зеленку;
                                    int KCP_temp = 0;
                                    if (int.TryParse(dgvAbitProfileList.Rows[1].Cells[cells.ColumnIndex].Value.ToString(), out KCP_temp))
                                    {
                                        for (int row_temp = startrow + KCP_temp; row_temp < dgvAbitProfileList.Rows.Count; row_temp++)
                                        {
                                            if ((dgvAbitProfileList.Rows[row_temp].Cells[cells.ColumnIndex].Style.BackColor != Color.LightGreen) &&
                                           (dgvAbitProfileList.Rows[row_temp].Cells[cells.ColumnIndex].Style.BackColor != Color.Yellow) &&
                                               (dgvAbitProfileList.Rows[row_temp].Cells[cells.ColumnIndex].Style.BackColor != Color.LightBlue))
                                            //if (dgvAbitProfileList.Rows[row_temp].Cells[cells.ColumnIndex].Style.BackColor == Color.White)
                                            {
                                                dgvAbitProfileList.Rows[row_temp].Cells[cells.ColumnIndex].Style.BackColor = Color.LightGreen;
                                                break;
                                            }
                                        }
                                    }
                                }
                                cells.Style.BackColor = Color.Yellow;
                            }
                        }
                    }
                }
            } 
        }

        private void btnPaint_Click(object sender, EventArgs e)
        {
            GridPaint();
            CopyTable();
            btnPaint.Enabled = false;
        }

        private void CopyTable()
        {
            int startrow = 2;
            int startcol = 1;
            for (int i = startrow; i < dgvAbitProfileList.Rows.Count; i++)
            {
                for (int j = startcol; j < dgvAbitProfileList.Columns.Count; j++)
                {
                    string value = dgvAbitProfileList.Rows[i].Cells[j].Value.ToString();
                    Guid PersId = Guid.Parse(value.Substring(0, value.IndexOf("_")));
                    dgvAbitProfileList.Rows[i].Cells[j].Value = PersonFIOList[PersonNumList.IndexOf(PersId)] + " (" + value.Substring(value.IndexOf("_") + 1) + ")";
                }
            }
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
    }
}
