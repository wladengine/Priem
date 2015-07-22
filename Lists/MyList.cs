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
    public partial class MyList : BookList
    {
        List<Guid> PersonList = new List<Guid>();
        List<string> PersonListFio = new List<string>();
        List<List<KeyValuePair<int, int>>> Coord = new List<List<KeyValuePair<int, int>>>();
        List<List<KeyValuePair<int, int>>> Coord_Save = new List<List<KeyValuePair<int, int>>>();
        List<KeyValuePair<int, KeyValuePair<int, int>>> DeleteList = new List<KeyValuePair<int, KeyValuePair<int, int>>>();
        Guid ErrorGuid = Guid.Empty;
        int startrow = 11;
        bool btnGreenIsClicked = false;

        int RowLicenceName, RowObrazProgramName, RowInnerObrazProgram, RowInnerProfile, RowEntryId, RowInnerEntryInEntryId, RowFacultyId, RowKCP, RowEge;

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
                    ComboServ.FillCombo(cbLicenseProgram, new List<KeyValuePair<string,string>>(), false, true);
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
            DataTable examTable = new DataTable();
            examTable.Columns.Add("Id");

            DataRow row_FacultyId = examTable.NewRow();
            DataRow row_LicProg = examTable.NewRow();
            DataRow row_ObrazProg = examTable.NewRow();
            DataRow row_InnerObrazProg = examTable.NewRow();
            DataRow row_InnerProfile = examTable.NewRow();
            DataRow row_EntryId = examTable.NewRow();
            DataRow row_InnerEntryInEntryId = examTable.NewRow();
            DataRow row_StudyForm = examTable.NewRow();
            DataRow row_StudyBasis = examTable.NewRow();
            DataRow row_Ege = examTable.NewRow();
            DataRow row_KCP = examTable.NewRow(); 

            DataColumn clm;

            PersonList = new List<Guid>();
            PersonListFio = new List<string>();
            Coord = new List<List<KeyValuePair<int, int>>>();

            NewWatch wc = new NewWatch();
            wc.Show();
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

            string index = "";
            int cnt = 1;

            DataTable tbl = MainClass.Bdc.GetDataSet(query).Tables[0];

            foreach (DataRow rw in tbl.Rows)
            {
                wc.SetText("Получение данных по учебным планам... (Обработано конкурсов: " + (cnt++).ToString() + "/" + tbl.Rows.Count + ")");
                clm = new DataColumn();
                index = Guid.NewGuid().ToString();//rwEntry.Field<int>("StudyFormId").ToString() + "_" + rwEntry.Field<int>("StudyBasisId").ToString() + "_" + rwEntry.Field<int>("LicenseProgramId").ToString() + "_" + rw_licProg.Field<int>("ObrazProgramId").ToString() + "_" + row_profile.Field<int>("ProfileId").ToString();
                clm.ColumnName = index;
                examTable.Columns.Add(clm);

                row_FacultyId[index] = rw.Field<int>("FacultyId");

                row_LicProg[index] = rw.Field<string>("LicenseProgramName");
                row_ObrazProg[index] = rw.Field<string>("ObrazProgramName");
                row_EntryId[index] = rw.Field<Guid>("EntryId");

                row_InnerEntryInEntryId[index] = rw.Field<Guid?>("InnerEntryInEntryId").ToString();
                row_InnerObrazProg[index] = rw.Field<string>("InnerObrazProgramName");
                row_InnerProfile[index] = rw.Field<string>("InnerProfileName");

                row_StudyForm[index] = (rw.Field<int>("StudyFormId") == 1) ? "Очная" : "Очно-заочная";
                row_StudyBasis[index] = (rw.Field<int>("StudyBasisId") == 1) ? "Бюджетная" : "Договорная";

                row_KCP[index] = rw.Field<int?>("InnerEntryInEntryKCP") ?? (rw.Field<int?>("extEntryKCP") ?? 0);

                row_Ege[index] = (!String.IsNullOrEmpty(rw.Field<int?>("EgeExamNameId").ToString())) ? rw.Field<int?>("EgeExamNameId") + "_" + rw.Field<string>("EgeName") + "(" + rw.Field<int?>("EgeMin") + ")" : "";
            }


            /*
            ///// Поиск по Направлениям в extEntry
            string query = @"Select distinct extEntry.LicenseProgramId, extEntry.FacultyId, extEntry.LicenseProgramName, extEntry.StudyBasisId, extEntry.StudyFormId
                                from ed.extEntry " + abitFilters + " order by StudyFormId, StudyBasisId, LicenseProgramName ";
            DataTable tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
            string index = "";
            int cnt = 1;
            foreach (DataRow rwEntry in tbl.Rows)
            {
                wc.SetText("Получение данных по учебным планам... (Обработано конкурсов: " + (cnt++).ToString() + "/" + tbl.Rows.Count + ")");
                ///// Поиск ОБРАЗОВАТЕЛЬНЫХ ПРОГРАММ 
                query = @"Select distinct extEntry.ObrazProgramId, extEntry.ObrazProgramName
                                from ed.extEntry " + abitFilters + 
                               " and LicenseProgramId=" + rwEntry.Field<int>("LicenseProgramId").ToString() + 
                               " and StudyBasisId=" + rwEntry.Field<int>("StudyBasisId").ToString() +
                               " and StudyFormId=" + rwEntry.Field<int>("StudyFormId").ToString() + 
                               " and IsSecond = 0";
                DataTable tbl_LicProg = MainClass.Bdc.GetDataSet(query).Tables[0];

                foreach (DataRow rw_licProg in tbl_LicProg.Rows)
                {
                    
                    ///// ДЛЯ КАЖДОЙ ОБРАЗОВАТЕЛЬНОЙ ПРОГРАММЫ ПОИСК ПРОФИЛЕЙ:
                    query = @"select distinct extEntry.Id, FacultyId, KCP, ProfileId, ProfileName 
                              from ed.extEntry" + abitFilters + 
                            " and LicenseProgramId=" + rwEntry.Field<int>("LicenseProgramId").ToString() +
                            " and ObrazProgramId=" + rw_licProg.Field<int>("ObrazProgramId").ToString() + 
                            " and ProfileId is not null and StudyBasisId=" + rwEntry.Field<int>("StudyBasisId").ToString() +
                            " and StudyFormId=" + rwEntry.Field<int>("StudyFormId").ToString() + 
                            " and IsSecond = 0";
                    DataTable tbl_ObrProgramProfile = MainClass.Bdc.GetDataSet(query).Tables[0];
                    /////  ЕСЛИ ЕСТЬ НЕНУЛЕВЫЕ ПРОФИЛИ (ПРОБЛЕМА С ИД СТОЛБЦА)
                    /////  НЕ ДОЛЖНО БЫТЬ ЗАГОЛОВКА СЛОБЦА, СТОЛБЕЦ = (НАПР/ОБРПРОГ/ПРОФ)
                    if (tbl_ObrProgramProfile.Rows.Count > 0)
                    {
                        foreach (DataRow row_profile in tbl_ObrProgramProfile.Rows)
                        {
                            clm = new DataColumn();
                            index = Guid.NewGuid().ToString();//rwEntry.Field<int>("StudyFormId").ToString() + "_" + rwEntry.Field<int>("StudyBasisId").ToString() + "_" + rwEntry.Field<int>("LicenseProgramId").ToString() + "_" + rw_licProg.Field<int>("ObrazProgramId").ToString() + "_" + row_profile.Field<int>("ProfileId").ToString();
                            clm.ColumnName = index;
                            examTable.Columns.Add(clm);
                            row_FacultyId[index] = rwEntry.Field<int>("FacultyId");
                            row_LicProg[index] = rwEntry.Field<string>("LicenseProgramName");
                            row_ObrazProg[index] = rw_licProg.Field<string>("ObrazProgramName") ;
                            row_Profile[index] = row_profile.Field<string>("ProfileName");
                            row_EntryId[index] = row_profile.Field<Guid>("Id");
                            row_InnerObrazProg[index] = "";
                            row_InnerEntryInEntryId[index] = "";
                            row_StudyForm[index] = (rwEntry.Field<int>("StudyFormId") == 1) ? "Очная" : "Очно-заочная";
                            row_StudyBasis[index] = (rwEntry.Field<int>("StudyBasisId") == 1) ? "Бюджетная" : "Договорная";
                            row_KCP[index] = row_profile.Field<int>("KCP");
                        }
                    }
                    /////  НЕНУЛЕВЫХ ПРОФИЛЕЙ НЕТ (ВОЗМОЖНО ЕСТЬ OBRAZ_PROGRAM_IN_ENTRY) 
                    else
                    {
                        //// нужно получить EntryId 
                        query = @"select distinct extEntry.Id, KCP 
                                  from ed.extEntry" + abitFilters + " and LicenseProgramId=" + rwEntry.Field<int>("LicenseProgramId").ToString() +
                                  " and ObrazProgramId=" + rw_licProg.Field<int>("ObrazProgramId").ToString() + 
                                  " and StudyBasisId=" + rwEntry.Field<int>("StudyBasisId").ToString() +
                                  " and StudyFormId=" + rwEntry.Field<int>("StudyFormId").ToString() + 
                                  " and IsSecond = 0";
                        DataSet ds = MainClass.Bdc.GetDataSet(query);
                        Guid EntryId = ds.Tables[0].Rows[0].Field<Guid>("Id");
                        int _KCP = ds.Tables[0].Rows[0].Field<int>("KCP");

                        /// поиск по EntryId В ОБРАЗОВАТЕЛЬНЫХ ПРОГРАММАХ
                        query = @"SELECT distinct InnerEntryInEntry.[Id] as Id, 
                                SP_ObrazProgram.Name as Name, SP_ObrazProgram.Id as ObrazProgramId, 
                                SP_Profile.Name as ProfileName,  ProfileId
                                KCP,
                                InnerEntryInEntry.EgeExamNameId , InnerEntryInEntry.EgeMin, EgeExamName.Name as EgeName, FacultyId
                                FROM [ed].[InnerEntryInEntry] 
                                inner join ed.SP_ObrazProgram on InnerEntryInEntry.ObrazProgramId = SP_ObrazProgram.Id
                                inner join ed.SP_Profile on SP_Profile.Id= ProfileId
                                left join ed.EgeExamName on EgeExamName.Id = InnerEntryInEntry.EgeExamNameId
                                where EntryId ='" + EntryId + @"'
                                order by ObrazProgramId";
                        DataTable tbl_ObrProgram = MainClass.Bdc.GetDataSet(query).Tables[0];
                        ///// приоритетов образ.программ нет
                        if (tbl_ObrProgram.Rows.Count == 0)
                        {
                            index = Guid.NewGuid().ToString();//rwEntry.Field<int>("StudyFormId").ToString() + "_" + rwEntry.Field<int>("StudyBasisId").ToString() + "_" + rwEntry.Field<int>("LicenseProgramId").ToString() + "_" + rw_licProg.Field<int>("ObrazProgramId").ToString() + "_0";
                            clm = new DataColumn();
                            clm.ColumnName = index;
                            examTable.Columns.Add(clm);
                            row_FacultyId[index] = rwEntry.Field<int>("FacultyId");
                            row_LicProg[index] = rwEntry.Field<string>("LicenseProgramName");
                            row_ObrazProg[index] = rw_licProg.Field<String>("ObrazProgramName");
                            row_Profile[index] = 
                            row_EntryId[index] = EntryId.ToString();
                            row_StudyForm[index] = (rwEntry.Field<int>("StudyFormId") == 1) ? "Очная" : "Очно-заочная";
                            row_StudyBasis[index] = (rwEntry.Field<int>("StudyBasisId") == 1) ? "Бюджетная" : "Договорная";
                            row_InnerEntryInEntryId[index] = "";
                            //
                            row_KCP[index] = _KCP;
                        }
                        else
                        {
                            ///// ПРИОРИТЕТЫ ОБРАЗ.ПРОГРАММ есть
                            foreach (DataRow rw_ObProg in tbl_ObrProgram.Rows)
                            {
                                clm = new DataColumn();
                                index = Guid.NewGuid().ToString();//rwEntry.Field<int>("StudyFormId").ToString() + "_" + rwEntry.Field<int>("StudyBasisId").ToString() + "_" + rwEntry.Field<int>("LicenseProgramId").ToString() + "_" + rw_ObProg.Field<int>("ObrazProgramId").ToString() + "_0";
                                clm.ColumnName = index;
                                examTable.Columns.Add(clm);
                                row_FacultyId[index] = rwEntry.Field<int>("FacultyId");
                                row_ObrazProg[index] = rw_ObProg.Field<String>("Name");
                                row_LicProg[index] = rwEntry.Field<string>("LicenseProgramName");
                                row_EntryId[index] = EntryId;
                                row_StudyForm[index] = (rwEntry.Field<int>("StudyFormId") == 1) ? "Очная" : "Очно-заочная";
                                row_StudyBasis[index] = (rwEntry.Field<int>("StudyBasisId") == 1) ? "Бюджетная" : "Договорная";
                                row_InnerEntryInEntryId[index] = rw_ObProg.Field<Guid>("Id").ToString();
                                row_Ege[index] = (!String.IsNullOrEmpty(rw_ObProg.Field<int?>("EgeExamNameId").ToString())) ? rw_ObProg.Field<int?>("EgeExamNameId") + "_" + rw_ObProg.Field<string>("EgeName") + "(" + rw_ObProg.Field<int?>("EgeMin") + ")" : "";
                                row_KCP[index] = rw_ObProg.Field<int>("KCP");
                            }
                        }
                    }
                }
                // ЗАКОНЧИЛСЯ ПОИСК ВНУТРИ ОБРАЗОВАТЕЛЬНОЙ ПРОГРАММЫ
            }
            */ 
            wc.SetText("Пересчет контрольных цифр приема...(0/" + (examTable.Columns.Count - 1).ToString() + ")");
            for (int i = 1; i < examTable.Columns.Count; i++)
            {
                wc.SetText("Пересчет контрольных цифр приема...("+(i+1)+"/" + (examTable.Columns.Count - 1).ToString() + ")");
                int kcp_new = int.Parse(row_KCP[i].ToString());
                if (!cbZeroWave.Checked)
                    if (!String.IsNullOrEmpty(row_InnerEntryInEntryId[i].ToString()))
                    {
                        row_KCP[i] = kcp_new - int.Parse(MainClass.Bdc.GetStringValue(@"
                                    select COUNT(extEntryView.Id) 
                                    from ed.extEntryView
                                    inner join ed.Abiturient on AbiturientId = Abiturient.Id
                                    where Abiturient.EntryId = '"+row_EntryId[i].ToString()+@"' and 
                                    Abiturient.InnerEntryInEntryId = '" + row_InnerEntryInEntryId[i].ToString() + @"'
                                    and Abiturient.CompetitionId NOT IN (12,11)"));
                    }
                    else
                    {
                        row_KCP[i] = kcp_new - int.Parse(MainClass.Bdc.GetStringValue(@"
                                    select COUNT(extEntryView.Id) 
                                    from ed.extEntryView
                                    inner join ed.Abiturient on AbiturientId = Abiturient.Id
                                    where Abiturient.EntryId = '" + row_EntryId[i].ToString() + @"'
                                    and Abiturient.CompetitionId NOT IN (12,11)"));
                    } 

            }
            examTable.Rows.Add(row_LicProg); //0
            RowLicenceName = 0;
            examTable.Rows.Add(row_ObrazProg);//1
            RowObrazProgramName = 1;
            examTable.Rows.Add(row_EntryId);//2
            RowEntryId = 2;

            examTable.Rows.Add(row_InnerEntryInEntryId); //3
            RowInnerEntryInEntryId = 3;

            examTable.Rows.Add(row_InnerObrazProg); //4
            RowInnerObrazProgram = 4;
            examTable.Rows.Add(row_InnerProfile); //5
            RowInnerProfile = 5;
            
            examTable.Rows.Add(row_StudyForm); //6
            examTable.Rows.Add(row_StudyBasis); //7
            examTable.Rows.Add(row_Ege); //8
            RowEge = 8;
            examTable.Rows.Add(row_KCP); //9

            RowKCP = 9;
            examTable.Rows.Add(row_FacultyId); //10
            RowFacultyId = 10;


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
            query = @"select " + toplist + @" Abiturient.Id, extPerson.PersonNum, Abiturient.PersonId, Abiturient.Priority, extAbitMarksSum.TotalSum, extPerson.FIO as FIO
            from ed.Abiturient
            inner join ed.extPerson on Abiturient.PersonId = extPerson.Id
            left join ed.extAbitMarksSum on extAbitMarksSum.Id = Abiturient.Id
            inner join ed." +Wave+ @" on "+Wave+ @".AbiturientId = Abiturient.Id
            inner join ed.extEntry on Abiturient.EntryId = extEntry.Id
            " + ((cbZeroWave.Checked) ? "inner join ed.extEntryView on extEntryView.AbiturientId = Abiturient.Id" : "") + 
            @"
            where Abiturient.EntryId=@EntryId and Abiturient.BackDoc = 0
            and Abiturient.CompetitionId NOT IN (12,11)
            --order by extAbitMarksSum.TotalSum desc
             order by  " + Wave + @".SortNum 
            ";
            wc.SetMax(examTable.Columns.Count - 1);
            wc.SetText("Получение данных по абитуриентам...(0/" + (examTable.Columns.Count - 1).ToString() + ")");
            for (int i = 1; i < examTable.Columns.Count; i++)
            {
                int j = 0;
                string _innerentryId = examTable.Rows[RowInnerEntryInEntryId][i].ToString();
                String EntryId = examTable.Rows[RowEntryId][i].ToString();
                DataSet ds = MainClass.Bdc.GetDataSet(query, new SortedList<string, object> { { "@EntryId", EntryId } });
                foreach (DataRow rw in ds.Tables[0].Rows)
                {
                    Guid _AbitId = rw.Field<Guid>("Id");
                    Guid _PersonId = rw.Field<Guid>("PersonId");
                    string FIO = rw.Field<string>("PersonNum") + "_" + rw.Field<string>("FIO");
                    if (!PersonList.Contains(_PersonId))
                    {
                        PersonList.Add(_PersonId);
                        PersonListFio.Add(FIO);
                    }

                    int _Priority = rw.Field<int?>("Priority") ?? 0;
                    int _innerprior = 0;
                    if (!String.IsNullOrEmpty(_innerentryId))
                    {
                        string query_obrazProgram = @"select InnerEntryInEntryPriority from ed.ApplicationDetails 
                              where ApplicationDetails.ApplicationId='" + _AbitId.ToString() + "' and InnerEntryInEntryId='" + _innerentryId + "'";
                        DataSet dsobrprog = MainClass.Bdc.GetDataSet(query_obrazProgram);
                        if (dsobrprog.Tables[0].Rows.Count > 0)
                            _innerprior = dsobrprog.Tables[0].Rows[0].Field<int?>("InnerEntryInEntryPriority") ?? 0;

                    }
                    String Temp_String = _Priority.ToString() + "_" + _innerprior.ToString() + "_" + _PersonId.ToString();
                    Temp_String += "_" + rw.Field<decimal?>("TotalSum").ToString();
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
                wc.SetText("Получение данных по абитуриентам...(Обработано конкурсов: " + i + "/" + (examTable.Columns.Count - 1).ToString() + ")");
            }
            for (int j = 0; j < RowList.Count; j++)
            {
                DataRow rw = RowList[j];
                examTable.Rows.Add(rw);
            }
            Coord_Save = Coord;

            DataView dv = new DataView(examTable);
            
            dgvAbitList.DataSource = dv;
            if (dgvAbitList.CurrentRow !=null )
            {
                if (dgvAbitList.CurrentRow.Index != RowFacultyId)
                    dgvAbitList.Rows[RowFacultyId].Visible = false;
                else
                {
                    dgvAbitList.CurrentCell = dgvAbitList.Rows[0].Cells[0];
                    dgvAbitList.Rows[RowFacultyId].Visible = false;
                }
            }
            dgvAbitList.Columns["Id"].Visible = false;
            dgvAbitList.ColumnHeadersVisible = false;
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
        private void cbStudyForm_SelectedIndexChanged(object sender, EventArgs e)
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
                            if (dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Style.BackColor == Color.LightGreen)
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
            /// Раскрашиваем профили
             /*
            for (int colindex = 1; colindex < dgvAbitList.Columns.Count; colindex++)
            {
                if (!String.IsNullOrEmpty(dgvAbitList.Rows[3].Cells[colindex].Value.ToString()))
                {
                    string query = @"select Id from ed.ProfileInObrazProgramInEntry where ObrazProgramInEntryId ='" + dgvAbitList.Rows[3].Cells[colindex].Value.ToString() + "'";
                    DataTable tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
                    if (tbl.Rows.Count < 2)
                        continue;
                    for (int j = 0; j < startrow - 1; j++)
                        dgvAbitList.Rows[j].Cells[colindex].Style.BackColor = Color.Azure;
                }
            }
            */
            ///
            for (int colindex = 1; colindex < dgvAbitList.Columns.Count; colindex++)
            {
                if (String.IsNullOrEmpty(dgvAbitList.Rows[RowInnerObrazProgram].Cells[colindex].Value.ToString()))
                {
                    dgvAbitList.Rows[RowInnerObrazProgram].Cells[colindex].Value = dgvAbitList.Rows[RowObrazProgramName].Cells[colindex].Value.ToString();
                    dgvAbitList.Rows[RowInnerProfile].Cells[colindex].Value = "(нет)";
                }
                int KCP = 0;
                int.TryParse(dgvAbitList.Rows[RowKCP].Cells[colindex].Value.ToString(), out KCP);

                for (int j = startrow; (j < KCP + startrow) && (j < dgvAbitList.Rows.Count); j++)
                {
                    if (String.IsNullOrEmpty(dgvAbitList.Rows[j].Cells[colindex].Value.ToString()))
                        break;
                    dgvAbitList.Rows[j].Cells[colindex].Style.BackColor = Color.LightGreen;
                }
            }
            for (int colindex = 1; colindex < dgvAbitList.Columns.Count; colindex++)
            {
                bool hasinnerprior = !String.IsNullOrEmpty((String)dgvAbitList.Rows[RowInnerEntryInEntryId].Cells[colindex].Value);
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

            // теперь Английские языки
            for (int colindex = 1; colindex < dgvAbitList.Columns.Count; colindex++)
            {
                if (String.IsNullOrEmpty(dgvAbitList.Rows[RowEge].Cells[colindex].Value.ToString()))
                    continue;

                string EgeExamNameId = dgvAbitList.Rows[RowEge].Cells[colindex].Value.ToString();
                EgeExamNameId = EgeExamNameId.Substring(0, EgeExamNameId.IndexOf("_"));
                string sEgeMin = dgvAbitList.Rows[RowEge].Cells[colindex].Value.ToString();
                sEgeMin = sEgeMin.Substring(sEgeMin.IndexOf("(") + 1);
                sEgeMin = sEgeMin.Substring(0, sEgeMin.IndexOf(")"));
                int EgeMin = int.Parse(sEgeMin);

                int KCP_temp = 0;
                if (int.TryParse(dgvAbitList.Rows[RowKCP].Cells[colindex].Value.ToString(), out KCP_temp))

                    for (int j = startrow; j < dgvAbitList.Rows.Count; j++)
                    {
                        string cellvalue = dgvAbitList.Rows[j].Cells[colindex].Value.ToString();
                        cellvalue = cellvalue.Substring(cellvalue.IndexOf('_') + 1);
                        cellvalue = cellvalue.Substring(cellvalue.IndexOf('_') + 1);
                        cellvalue = cellvalue.Substring(0, cellvalue.IndexOf('_'));
                        int EgeAbitValue = (int?)MainClass.Bdc.GetValue("select Value from ed.extEgeMark where PersonId = '" + cellvalue + "' and EgeExamNameId=" + EgeExamNameId + " and FBSStatusId=1") ?? 0;
                        if (EgeAbitValue < EgeMin)
                        {
                            if ((dgvAbitList.Rows[j].Cells[colindex].Style.BackColor == Color.LightGreen) || (dgvAbitList.Rows[j].Cells[colindex].Style.BackColor == Color.LightBlue))
                            {
                                // сдвинуть зеленку;
                                for (int row_temp = startrow + KCP_temp; row_temp < dgvAbitList.Rows.Count; row_temp++)
                                {
                                    if (String.IsNullOrEmpty(dgvAbitList.Rows[row_temp].Cells[colindex].Value.ToString()))
                                        break;
                                    if (dgvAbitList.Rows[row_temp].Cells[colindex].Style.BackColor == Color.Empty)
                                    {
                                        dgvAbitList.Rows[row_temp].Cells[colindex].Style.BackColor = Color.LightGreen;
                                        break;
                                    }
                                }
                            }
                            dgvAbitList.Rows[j].Cells[colindex].Style.BackColor = Color.Thistle;
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
                    MessageBox.Show("Цикл перерасчета приоритетов был произведен " + dgvAbitList.Rows.Count + " раз, наверно возникла какая-то ошибка. Придется прекратить перерасчет приоритетов. Проблемный Guid: " + ErrorGuid.ToString(), "Вы знаете,..", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                }
                for (int colindex = 1; colindex < dgvAbitList.Columns.Count; colindex++)
                {
                    int KCP = 0;
                    if (int.TryParse(dgvAbitList.Rows[RowKCP].Cells[colindex].Value.ToString(), out KCP))
                    { }


                    bool hasinnerprior = !String.IsNullOrEmpty((String)dgvAbitList.Rows[RowInnerEntryInEntryId].Cells[colindex].Value);
                    for (int j = startrow; (j < dgvAbitList.Rows.Count); j++)
                    {
                        if (dgvAbitList.Rows[j].Cells[colindex].Style.BackColor == Color.Empty)
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
                                // по всем координатам key;value = столбец; строка
                                foreach (KeyValuePair<int, int> kvp in Coord[index])
                                {

                                    // если это та же строка и тот же столбец
                                    if ((kvp.Value + startrow == j) && (kvp.Key == colindex))
                                    {
                                        continue;
                                    }

                                    int KCP_temp = 0;
                                    if (int.TryParse(dgvAbitList.Rows[RowKCP].Cells[kvp.Key].Value.ToString(), out KCP_temp))
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
                                        if (innerprior_temp >= innerprior)
                                        {
                                            if (innerprior_temp == innerprior)
                                            {
                                                string FIO = PersonListFio[index];
                                                MessageBox.Show(this, "Вы знаете, у абитуриента: " + FIO + " существуют повторяющиеся приоритеты", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                            }
                                            bool isGreen = false;
                                            if (dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Style.BackColor == Color.LightGreen)
                                            {
                                                isGreen = true;
                                            }

                                            // под вопросом #подумать
                                            if ((dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Style.BackColor == Color.LightGreen) ||
                                                (dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Style.BackColor == Color.LightBlue) ||
                                                (dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Style.BackColor == Color.Empty))
                                                dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Style.BackColor = Color.Yellow;
                                            if (innerprior_temp == innerprior)
                                            {
                                                dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Style.BackColor = Color.Red;
                                            }
                                            DeleteList.Add(new KeyValuePair<int, KeyValuePair<int, int>>(index, kvp));
                                            if (isGreen)
                                            {
                                                for (int row_temp = startrow + KCP_temp; row_temp < dgvAbitList.Rows.Count; row_temp++)
                                                {
                                                    if (String.IsNullOrEmpty(dgvAbitList.Rows[row_temp].Cells[kvp.Key].Value.ToString()))
                                                        break;
                                                    if (dgvAbitList.Rows[row_temp].Cells[kvp.Key].Style.BackColor == Color.Empty)
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
                                        if (prior_temp >= prior)
                                        {
                                            if (prior_temp == prior)
                                            {
                                                string FIO = PersonListFio[index];
                                                MessageBox.Show(this, "Вы знаете, у абитуриента: " + FIO + " существуют повторяющиеся приоритеты", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                            }
                                            bool isGreen = false;
                                            if (dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Style.BackColor == Color.LightGreen)
                                            {
                                                isGreen = true;
                                            }
                                            dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Style.BackColor = Color.Yellow;
                                            if (prior_temp == prior)
                                            {
                                                dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Style.BackColor = Color.Red;
                                            }
                                            //Coord[index].Remove(kvp);
                                            DeleteList.Add(new KeyValuePair<int, KeyValuePair<int, int>>(index, kvp));
                                            if (isGreen)
                                            {
                                                for (int row_temp = startrow + KCP_temp; row_temp < dgvAbitList.Rows.Count; row_temp++)
                                                {
                                                    if (String.IsNullOrEmpty(dgvAbitList.Rows[row_temp].Cells[kvp.Key].Value.ToString()))
                                                        break;
                                                    if (dgvAbitList.Rows[row_temp].Cells[kvp.Key].Style.BackColor == Color.Empty)
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
            /*
            for (int j = 1; j < dgvAbitList.Columns.Count; j++)
            {
                string value = dgvAbitList.Rows[6].Cells[j].Value.ToString();
                if (!String.IsNullOrEmpty(value))
                    dgvAbitList.Rows[6].Cells[j].Value = value.Substring(value.IndexOf("_") + 1);
            }
             */
            for (int i = startrow; i < dgvAbitList.Rows.Count; i++)
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
                    dcell.Value = FIO + " (" + _prior + ", " + _prior_inner + ", " + ball + ")";
                }
            }
             dgvAbitList.Rows[RowEntryId].Visible = false;
             dgvAbitList.Rows[RowInnerEntryInEntryId].Visible = false;
             dgvAbitList.Rows[RowObrazProgramName].Visible = false;
        }

        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if (dgvAbitList.Rows.Count > startrow)
            {
                DataTable tbl = ((DataView)dgvAbitList.DataSource).Table.Copy();

                string sheetName = "export";


                if (tbl.Columns.Contains("Id"))
                {
                    tbl.Columns.Remove("Id");
                }
                tbl.Rows[RowEntryId].Delete();
                tbl.Rows[RowInnerEntryInEntryId].Delete();
                tbl.Rows[RowObrazProgramName].Delete();



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
                                Color clr = dgvAbitList.Rows[rowindex + 2].Cells[colindex + 1].Style.BackColor;
                                if (clr != Color.Empty)
                                    Range3.Interior.Color = dgvAbitList.Rows[rowindex + 2].Cells[colindex + 1].Style.BackColor;
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

        private void dgvAbitList_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < startrow)
                return;
            if (e.ColumnIndex < 1)
                return;

            string FIO = dgvAbitList.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Substring(0, dgvAbitList.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().IndexOf('(') - 1);
            int index = PersonListFio.IndexOf(FIO);

            MainClass.OpenCardPerson(PersonList[index].ToString(), this, dgvAbitList.CurrentRow.Index);
        }

        private void dgvAbitList_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex < 1)
                return;
            // есть ли профили
            if (e.RowIndex < startrow)
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

                if (e.Button == MouseButtons.Right)
                {
                    dgvAbitList.CurrentCell = dgvAbitList.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    ContextMenu m = new ContextMenu();
                    m.MenuItems.Add(new MenuItem("Перейти к зеленой позиции", new EventHandler(this.ContextMenuToGreen_OnClick)));
                    m.MenuItems.Add(new MenuItem("Открыть карточку абитуриента", new EventHandler(this.ContextMenuOpenCard_OnClick)));

                    Point pCell = dgvAbitList.GetCellDisplayRectangle(dgvAbitList.CurrentCell.ColumnIndex, dgvAbitList.CurrentCell.RowIndex, true).Location;
                    Point pGrid = dgvAbitList.Location;
                    new Point(pCell.X + pGrid.X, pCell.Y + pGrid.Y + dgvAbitList.CurrentRow.Height);

                    m.Show(dgvAbitList, new Point(pCell.X + pGrid.X, pCell.Y + dgvAbitList.CurrentRow.Height));
                }
        }
        private void ContextMenuToGreen_OnClick(object sender, EventArgs e)
        {

            string FIO = dgvAbitList.CurrentCell.Value.ToString().Substring(0, dgvAbitList.CurrentCell.Value.ToString().IndexOf('(') - 1);
            int index = PersonListFio.IndexOf(FIO);
            if (index > -1)
                foreach (KeyValuePair<int, int> kvp in Coord_Save[index])
                {
                    if (dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key].Style.BackColor == Color.LightGreen)
                    {
                        dgvAbitList.CurrentCell = dgvAbitList.Rows[kvp.Value + startrow].Cells[kvp.Key];
                    }
                }
        }
        private void ContextMenuOpenCard_OnClick(object sender, EventArgs e)
        {
            string FIO = dgvAbitList.CurrentCell.Value.ToString().Substring(0, dgvAbitList.CurrentCell.Value.ToString().IndexOf('(') - 1);
            int index = PersonListFio.IndexOf(FIO);

            MainClass.OpenCardPerson(PersonList[index].ToString(), this, dgvAbitList.CurrentRow.Index);
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

        private void tbAbitsTop_MouseClick(object sender, MouseEventArgs e)
        {
            rbAbitsTop.Checked = true;
        }

        private void btn_GreenList_Click(object sender, EventArgs e)
        {
            int startcol = 1;
            NewWatch wc = new NewWatch();
            wc.Show();
            wc.SetText("Удаление старых данных...");
            MainClass.Bdc.ExecuteQuery(@"  delete from ed.AbiturientGREEN
                                           where AbiturientId In (select qAbiturient.Id from ed.qAbiturient inner join ed.extEntry on extEntry.Id = EntryId " + GetAbitFilterString() + ")");
            wc.SetText("Добавление новых данных...");
            wc.SetMax(dgvAbitList.Columns.Count);
            for (int clmn = startcol; clmn < dgvAbitList.Columns.Count; clmn++)
            {
                // 0 LicenseProgramName
                // 1 ObrazProgramName
                // 2 EntryId
                // 3 obrazprogramInEntryId
                // 4 форма
                // 5 основа
                // 6 кцп
                // 7 абитуриентик
                string InnerEntryInEntryId = dgvAbitList.Rows[RowInnerEntryInEntryId].Cells[clmn].Value.ToString();
                string EntryId = dgvAbitList.Rows[RowEntryId].Cells[clmn].Value.ToString();
                string PersonId = "";
                string AbitId = "";
                string NumFio = "";
                string query = @"select Abiturient.Id from ed.Abiturient where BackDoc=0 and NotEnabled = 0 and  EntryId ='" + EntryId + "' and PersonId='";

                // человеки
                List<string> AbitIdList = new List<string>();

                for (int rowindex = startrow; rowindex < dgvAbitList.Rows.Count; rowindex++)
                {
                    string value = dgvAbitList.Rows[rowindex].Cells[clmn].Value.ToString();
                    if (String.IsNullOrEmpty(value))
                        break;
                    if (dgvAbitList.Rows[rowindex].Cells[clmn].Style.BackColor == Color.Empty)
                        break;

                    if (dgvAbitList.Rows[rowindex].Cells[clmn].Style.BackColor == Color.LightGreen)  
                    {
                        NumFio = value.Substring(0, value.IndexOf("(") - 1);
                        int personIndexInList = PersonListFio.IndexOf(NumFio);
                        if (personIndexInList > -1)
                        {
                            PersonId = PersonList[personIndexInList].ToString();
                            AbitId = MainClass.Bdc.GetDataSet(query + PersonId + "'").Tables[0].Rows[0].Field<Guid>("Id").ToString();
                            if (!String.IsNullOrEmpty(AbitId))
                            {
                                // добавляем ID абитуриента в список счастливых лиц
                                AbitIdList.Add(AbitId);
                                if (string.IsNullOrEmpty(InnerEntryInEntryId))
                                    MainClass.Bdc.ExecuteQuery("Insert into  ed.AbiturientGreen (AbiturientId) Values ('" + AbitId + "')");
                                else
                                    MainClass.Bdc.ExecuteQuery("Insert into  ed.AbiturientGreen  (AbiturientId, InnerEntryInEntryId) Values ('" + AbitId + "', '" + InnerEntryInEntryId + "')");
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
                wc.SetText("Добавление новых данных: Обработано конкурсов " + clmn + "/" + (dgvAbitList.Columns.Count - 1) + "...");
            }
            wc.Close();
            btnGreenIsClicked = true;
            MessageBox.Show(this, "Done", "", MessageBoxButtons.OK);
        }

        private void dgvAbitList_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        { 
        }

        private void cbZeroWave_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void FillGrid(bool update)
        {
            if (update)
                FillGrid(GetAbitFilterString());

            if (FacultyId != null)
            {
                foreach (DataGridViewColumn clm in dgvAbitList.Columns)
                {
                    if (dgvAbitList.Rows[RowFacultyId].Cells[clm.Index].Value.ToString() != FacultyId.Value.ToString())
                    {
                        clm.Visible = false;
                    }
                    else
                        clm.Visible = true;
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
                        if (dgvAbitList.Rows[RowLicenceName].Cells[clm.Index].Value.ToString() != ((KeyValuePair<string,string>)cbLicenseProgram.SelectedItem).Value)
                        {
                            clm.Visible = false;
                        }
                        else
                            clm.Visible = true;
                    }
                }

            }
            if (dgvAbitList.Columns.Contains("Id"))
                dgvAbitList.Columns["Id"].Visible = false;
            if (update)
                btn_GreenList.Enabled = true;
        }
    }
}