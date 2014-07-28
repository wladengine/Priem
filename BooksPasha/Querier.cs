using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.IO;

namespace Priem
{
    public partial class Querier : Form
    {
        public Querier()
        {
            InitializeComponent();
        }

        private void btnPayDataEntryLoadCSV_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "CSV files|*.csv";
            if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return;

            bool isSecond = false;
            if (MessageBox.Show("IsSecond?", "IsSecond?", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                isSecond = true;

            
            using (PriemEntities context = new PriemEntities())
            {
                var entryList = context.qEntry.Where(x => x.StudyBasisId == 2 && x.StudyLevelGroupId == MainClass.studyLevelGroupId && x.IsSecond == isSecond)
                    .Select(x => new
                    {
                        x.Id,
                        x.LicenseProgramCode,
                        x.LicenseProgramName,
                        x.ObrazProgramCrypt,
                        x.ObrazProgramName,
                        x.ProfileName,
                        x.StudyFormName
                    }).ToList()
                    .Select(x => new
                    {
                        x.Id,
                        LP = x.LicenseProgramCode + " " + x.LicenseProgramName,
                        OP = x.ObrazProgramCrypt + " " + x.ObrazProgramName,
                        x.ObrazProgramName,
                        x.ProfileName,
                        x.StudyFormName
                    }).OrderBy(x => x.LP).ThenBy(x => x.OP);

                string BigStr = string.Empty;
                using (StreamReader sr = new StreamReader(ofd.FileName, Encoding.GetEncoding(1251)))
                {
                    BigStr = sr.ReadToEnd();
                }

                //!!!ЭТОТ КУСОК КОДА ОТКРЫВАЕТ ЛЮБОЙ ЗАДРОЧЕННЫЙ ТОЧКАМИ С ЗАПЯТЫМИ ФАЙЛ!
                //!!!КОД ХРАНИТЬ, ОН ЕЩЁ ПОМОЖЕТ!
                while (BigStr.Length > 0)
                {
                    string str = string.Empty;// = //BigStr.Substring(
                    for (int i = 0; (i < 7 && BigStr.Length > 0); i++)
                    {
                        int ind = BigStr.IndexOf(';');
                        string substr = BigStr.Substring(0, ind + 1);
                        str += substr;
                        BigStr = BigStr.Length > ind ? BigStr.Substring(ind + 1) : "";
                        if (substr.StartsWith("\"", StringComparison.OrdinalIgnoreCase) && !substr.EndsWith("\";", StringComparison.OrdinalIgnoreCase))
                        {
                            while (!substr.EndsWith("\";", StringComparison.OrdinalIgnoreCase))
                            {
                                ind = BigStr.IndexOf(';');
                                substr = BigStr.Substring(0, ind + 1);
                                str += substr;
                                BigStr = BigStr.Length > ind ? BigStr.Substring(ind + 1) : "";
                                if (string.IsNullOrEmpty(BigStr))
                                    break;
                            }
                        }
                    }

                    string[] splitted = str.Split(';');
                    if (splitted.Count() < 8)
                        continue;

                    string[] tmpSplitted = new string[8];
                    int index = 0;
                    foreach (string s in splitted)
                    {
                        if (s.StartsWith("\"", StringComparison.OrdinalIgnoreCase) && !s.EndsWith("\"", StringComparison.OrdinalIgnoreCase))
                        {
                            tmpSplitted[index] += s + ";";//доставляем пропущенную почём зря точку с запятой
                            continue;
                        }
                        if (s.EndsWith("\"", StringComparison.OrdinalIgnoreCase))
                        {
                            tmpSplitted[index] += s;
                            index++;
                            continue;
                        }
                        if (!s.StartsWith("\"", StringComparison.OrdinalIgnoreCase))
                        {
                            tmpSplitted[index] += s;
                            index++;
                            continue;
                        }
                    }

                    splitted = tmpSplitted;

                    

                    PayDataEntry pd = new PayDataEntry();
                    string LP = splitted[0].Replace("\r\n", "").Trim().Trim('"');
                    string OP = splitted[1].Trim().Trim('"');
                    string ProfName = splitted[2].Trim().Trim('"');
                    string SF = splitted[3].Trim().Trim('"');

                    Guid? entryId = entryList.Where(x => x.LP.IndexOf(LP, StringComparison.OrdinalIgnoreCase) == 0 
                        && (x.OP.IndexOf(OP, StringComparison.OrdinalIgnoreCase) == 0 || x.ObrazProgramName.IndexOf(OP, StringComparison.OrdinalIgnoreCase) == 0)
                        && x.ProfileName.IndexOf(ProfName, StringComparison.OrdinalIgnoreCase) == 0 && x.StudyFormName.IndexOf(SF, StringComparison.OrdinalIgnoreCase) == 0).Select(x => x.Id).FirstOrDefault();

                    if (!entryId.HasValue || entryId == Guid.Empty)
                    {
                        MessageBox.Show("LP: " + LP + "\nOP: " + OP + "\nProfile: " + ProfName + "\nSF: " + SF + "\nFAIL - NO GUID");
                        continue;
                    }

                    bool bIns = true;
                    if (context.PayDataEntry.Where(x => x.EntryId == entryId).Count() == 1)
                    {
                        pd = context.PayDataEntry.Where(x => x.EntryId == entryId).First();
                        bIns = false;
                    }

                    int ProrectorId = 0;
                    string sProrectorName = splitted[4];
                    if (context.Prorektor.Where(x => x.NameFull == sProrectorName).Count() == 0)
                    {
                        Prorektor pr = new Prorektor();
                        pr.Name = sProrectorName;
                        pr.NameFull = sProrectorName;
                        pr.NumberDov = splitted[6];
                        ProrectorId = context.Prorektor.Select(x => x.Id).DefaultIfEmpty(0).Max() + 1;
                        pr.Id = ProrectorId;

                        context.Prorektor.AddObject(pr);
                        context.SaveChanges();
                    }
                    else
                    {
                        ProrectorId = context.Prorektor.Where(x => x.NameFull == sProrectorName).Select(x => x.Id).First();
                    }

                    if (bIns)
                        context.PayDataEntry_Insert(entryId.Value, "", "", "", "", "", ProrectorId, "бакалавр", "4 года", new DateTime(2013, 9, 1), new DateTime(2017, 8, 31), splitted[5]);
                    else
                        context.PayDataEntry_Update(entryId.Value, "", "", "", "", "", ProrectorId, "бакалавр", "4 года", new DateTime(2013, 9, 1), new DateTime(2017, 8, 31), splitted[5]);

                }
            }

            MessageBox.Show("DONE!");
        }

        private void btnImportExamSpecAspirant_Click(object sender, EventArgs e)
        {
            //some code here
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "CSV files|*.csv";
            if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return;
            string BigStr = string.Empty;
            using (StreamReader sr = new StreamReader(ofd.FileName, Encoding.GetEncoding(1251)))
            {
                BigStr = sr.ReadToEnd();
            }

            List<string> NumMarkList = new List<string>();
            while (BigStr.Length > 0)
            {
                int ind = BigStr.IndexOf("\r\n");
                string substr ="";
                if (ind > -1)
                    substr = BigStr.Substring(0, ind);
                else
                    substr = BigStr;

                if (substr.Length <= 0)
                    break;
                NumMarkList.Add(substr);
                if (ind == -1)
                {
                    BigStr  = "";break; 
                }
                BigStr = BigStr.Length > ind ? BigStr.Substring(ind + 2) : "";
            }
            //string nummarkstring = NumMarkList[0];
            foreach (string nummarkstring in NumMarkList)
            {
                string[] splitted = nummarkstring.Split(';');
                /*if (splitted.Count() < 2)
                    continue;*/
                string query = @"select Id, EntryId from ed.Abiturient where RegNum = '" + splitted[0]+"'";
                DataTable tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
                if (tbl.Rows.Count == 1)
                {
                    DataRow rw = tbl.Rows[0];
                    string abitId = rw.Field<Guid>("Id").ToString();
                    string EntryId = rw.Field<Guid>("EntryId").ToString();

                    query = @"select Id from ed.extExamInEntry where EntryId = '" + EntryId + "' and IsProfil = 1 and ExamName not in ('Философия (асп)', 'Иностранный язык (асп)')";
                    tbl = MainClass.Bdc.GetDataSet(query).Tables[0];
                    if (tbl.Rows.Count == 1)
                    {
                        rw = tbl.Rows[0];
                        string MarkInEntryId = rw.Field<int>("Id").ToString();
                        query = @"update ed.[Mark]  set FiveGradeValue = " + splitted[1] + " where AbiturientId = '" + abitId + "' and ExamInEntryId = " + MarkInEntryId;
                        MainClass.Bdc.ExecuteQuery(query);
                    }
                }
            } 
             
        }
    }
}
