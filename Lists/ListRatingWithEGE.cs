using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using EducServLib;

namespace Priem.Lists
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
                                   select new { EgeToEx.EgeExamNameId }).Distinct().ToList();


            }
        }
    }
}
