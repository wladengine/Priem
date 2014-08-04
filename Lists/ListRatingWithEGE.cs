using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using EducServLib;

namespace Priem
{
    public partial class ListRatingWithEGE : Form
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
        public Guid? ProfileId
        {
            get
            {
                string prId = ComboServ.GetComboId(cbProfile);
                if (string.IsNullOrEmpty(prId))
                    return null;
                else
                    return new Guid(prId);
            }
            set
            {
                if (value == null)
                    ComboServ.SetComboId(cbProfile, (string)null);
                else
                    ComboServ.SetComboId(cbProfile, value.ToString());
            }
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
        public bool IsSecond
        {
            get { return chbIsSecond.Checked; }
            set { chbIsSecond.Checked = value; }
        }
        public bool IsReduced
        {
            get { return chbIsReduced.Checked; }
            set { chbIsReduced.Checked = value; }
        }
        public bool IsParallel
        {
            get { return chbIsParallel.Checked; }
            set { chbIsParallel.Checked = value; }
        }

        public ListRatingWithEGE()
        {
            InitializeComponent();
            ExtraInit();
        }

        protected void ExtraInit()
        {
            ComboServ.FillCombo(cbFaculty, HelpClass.GetComboListByTable("ed.qFaculty", "ORDER BY Acronym"), false, false);
            ComboServ.FillCombo(cbStudyBasis, HelpClass.GetComboListByTable("ed.StudyBasis", "ORDER BY Name"), false, false);

            cbStudyBasis.SelectedIndex = 0;
            FillStudyForm();
            FillLicenseProgram();
            FillObrazProgram();
            FillProfile();

            InitHandlers();
        }

        private void FillStudyForm()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = MainClass.GetEntry(context).Where(c => c.FacultyId == FacultyId).Where(c=>c.StudyBasisId == StudyBasisId);
                
                ent = ent.Where(c => c.IsSecond == IsSecond && c.IsReduced == IsReduced && c.IsParallel == IsParallel);

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.StudyFormId.ToString(), u.StudyFormName)).Distinct().ToList();

                ComboServ.FillCombo(cbStudyForm, lst, false, false);                
            }
        }
        private void FillLicenseProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = MainClass.GetEntry(context).Where(c => c.FacultyId == FacultyId);

                ent = ent.Where(c => c.IsSecond == IsSecond && c.IsReduced == IsReduced && c.IsParallel == IsParallel);

                if (StudyBasisId != null)
                    ent = ent.Where(c => c.StudyBasisId == StudyBasisId);
                if (StudyFormId != null)
                    ent = ent.Where(c => c.StudyFormId == StudyFormId);

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.LicenseProgramId.ToString(), u.LicenseProgramName)).Distinct().ToList();

                ComboServ.FillCombo(cbLicenseProgram, lst, false, false);                
            }
        }
        private void FillObrazProgram()
        {
            using (PriemEntities context = new PriemEntities())
            {
                var ent = MainClass.GetEntry(context).Where(c => c.FacultyId == FacultyId);

                ent = ent.Where(c => c.IsSecond == IsSecond && c.IsReduced == IsReduced && c.IsParallel == IsParallel);

                if (StudyBasisId != null)
                    ent = ent.Where(c => c.StudyBasisId == StudyBasisId);
                if (StudyFormId != null)
                    ent = ent.Where(c => c.StudyFormId == StudyFormId);
                if (LicenseProgramId != null)
                    ent = ent.Where(c => c.LicenseProgramId == LicenseProgramId);

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.ObrazProgramId.ToString(), u.ObrazProgramName + ' ' + u.ObrazProgramCrypt)).Distinct().ToList();

                ComboServ.FillCombo(cbObrazProgram, lst, false, false);
            }
        }
        private void FillProfile()
        {
            using (PriemEntities context = new PriemEntities())
            {
                if (ObrazProgramId == null)
                {
                    ComboServ.FillCombo(cbProfile, new List<KeyValuePair<string, string>>(), false, false);
                    cbProfile.Enabled = false;
                    return;
                }

                var ent = MainClass.GetEntry(context).Where(c => c.FacultyId == FacultyId).Where(c => c.ProfileId != null);

                ent = ent.Where(c => c.IsSecond == IsSecond && c.IsReduced == IsReduced && c.IsParallel == IsParallel);

                if (StudyBasisId != null)
                    ent = ent.Where(c => c.StudyBasisId == StudyBasisId);
                if (StudyFormId != null)
                    ent = ent.Where(c => c.StudyFormId == StudyFormId);
                if (LicenseProgramId != null)
                    ent = ent.Where(c => c.LicenseProgramId == LicenseProgramId);
                if (ObrazProgramId != null)
                    ent = ent.Where(c => c.ObrazProgramId == ObrazProgramId);

                List<KeyValuePair<string, string>> lst = ent.ToList().Select(u => new KeyValuePair<string, string>(u.ProfileId.ToString(), u.ProfileName)).Distinct().ToList();

                if (lst.Count() > 0)
                {
                    ComboServ.FillCombo(cbProfile, lst, false, false);
                    cbProfile.Enabled = true;
                }
                else
                {
                    ComboServ.FillCombo(cbProfile, new List<KeyValuePair<string, string>>(), false, false);
                    cbProfile.Enabled = false;
                }              
            }
        }
        
        //инициализация обработчиков мегакомбов
        public void InitHandlers()
        {
            cbFaculty.SelectedIndexChanged += cbFaculty_SelectedIndexChanged;
            cbStudyForm.SelectedIndexChanged += cbStudyForm_SelectedIndexChanged;
            cbStudyBasis.SelectedIndexChanged += cbStudyBasis_SelectedIndexChanged;
            cbLicenseProgram.SelectedIndexChanged += cbLicenseProgram_SelectedIndexChanged;
            cbObrazProgram.SelectedIndexChanged += cbObrazProgram_SelectedIndexChanged;

            chbIsParallel.CheckedChanged += chbIsParallel_CheckedChanged;
            chbIsReduced.CheckedChanged += chbIsReduced_CheckedChanged;
            chbIsSecond.CheckedChanged += chbIsSecond_CheckedChanged;
        }

        private void chbIsReduced_CheckedChanged(object sender, EventArgs e)
        {
            FillStudyForm();
        }
        private void chbIsParallel_CheckedChanged(object sender, EventArgs e)
        {
            FillStudyForm();
        }
        private void chbIsSecond_CheckedChanged(object sender, EventArgs e)
        {
            FillStudyForm();
        }

        void cbFaculty_SelectedIndexChanged(object sender, EventArgs e)
        {            
            FillStudyForm();
        }
        void cbStudyBasis_SelectedIndexChanged(object sender, EventArgs e)
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
            FillProfile();
        }       

        private void btnStartCount_Click(object sender, EventArgs e)
        {
            using (PriemEntities context = new PriemEntities())
            {
                var entryList = (from ent in MainClass.GetEntry(context)
                             where ent.IsSecond == IsSecond && ent.IsParallel == IsParallel && ent.IsReduced == IsReduced
                             && ent.LicenseProgramId == LicenseProgramId
                             && ent.ObrazProgramId == ObrazProgramId
                             && (ProfileId == null ? ent.ProfileId == null : ent.ProfileId == ProfileId)
                             && ent.StudyFormId == StudyFormId
                             && ent.StudyBasisId == StudyBasisId
                             select ent.Id).ToList();

                if (entryList.Count() == 0)
                    return;
                if (entryList.Count() > 1)
                    return;

                Guid EntryId = entryList[0];

                var examInEntry = (from ExInEnt in context.ExamInEntry
                                   join EgeToEx in context.EgeToExam on ExInEnt.ExamId equals EgeToEx.ExamId
                                   where ExInEnt.EntryId == EntryId
                                   select EgeToEx.EgeExamNameId).Distinct().ToList();

                string EgeSelect = "";
                foreach (var exInEntExNameId in examInEntry)
                {
                    var exName = context.EgeExamName.Where(x => x.Id == exInEntExNameId).FirstOrDefault();
                    if (exName == null)
                        continue;

                    EgeSelect +=  string.Format("\n(SELECT MAX(extEgeMark.Value) FROM ed.extEgeMark WHERE extEgeMark.PersonId = extAbit.PersonId " +
                        " AND extEgeMark.FBSStatusId IN (1, 4) AND extEgeMark.EgeExamNameId = {0}) as [ЕГЭ {1}], ", exInEntExNameId, exName.Name);
                }

                string query = "SELECT extAbit.RegNum AS [Рег номер], extPerson.FIO AS [ФИО], extAbit.Priority AS [Приоритет], " + EgeSelect +
                    @"extAbitMarksSum.TotalSum AS [Сумма баллов], 
(CASE WHEN EXISTS 
	(
		SELECT * 
		FROM ed.Abiturient 
		INNER JOIN ed._FirstWaveGreen _F ON _F.AbiturientId = Abiturient.Id 
		WHERE Abiturient.PersonId = extAbit.PersonId AND Abiturient.Priority < extAbit.Priority AND Abiturient.Id <> extAbit.Id
	)
	THEN 1 ELSE 0 END
) AS [Проходит по старшему приоритету]
FROM ed.Abiturient AS extAbit
INNER JOIN ed.extEntry ON extEntry.Id = extAbit.EntryId
INNER JOIN ed.extPerson ON extPerson.Id = extAbit.PersonId
INNER JOIN ed.extAbitMarksSum ON extAbit.Id = extAbitMarksSum.Id
INNER JOIN ed._FirstWaveGreen ON extAbit.Id = _FirstWaveGreen.AbiturientId
INNER JOIN ed._FirstWave FW ON FW.AbiturientId = extAbit.Id
";
                string where = @" WHERE extEntry.Id = @EntryId ";
                string orderby = " ORDER BY FW.SortNum";

                DataTable tbl = MainClass.Bdc.GetDataSet(query + where + orderby, new SortedList<string, object>() { { "@EntryId", EntryId } }).Tables[0];

                dgvData.DataSource = tbl;
            }
        }

        private void btnSaveToExcel_Click(object sender, EventArgs e)
        {
            if (dgvData.Rows.Count > 0)
                PrintClass.PrintAllToExcel2007((DataTable)dgvData.DataSource, "upload");
        }
    }
}
