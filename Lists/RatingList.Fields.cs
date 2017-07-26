using EducServLib;
using PriemLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Priem
{
    public partial class RatingList
    {
        public int? StudyLevelGroupId
        {
            get { return ComboServ.GetComboIdInt(cbStudyLevelGroup); }
            set { ComboServ.SetComboId(cbStudyLevelGroup, value); }
        }
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
        public int? ProfileId
        {
            get
            {
                return ComboServ.GetComboIdInt(cbProfile);
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
        public bool IsCel
        {
            get { return chbCel.Checked; }
            set { chbCel.Checked = value; }
        }
        public bool IsQuota
        {
            get { return chbIsQuota.Checked; }
            set { chbIsQuota.Checked = value; }
        }
        public Guid? EntryId
        {
            get
            {
                try
                {
                    using (PriemEntities context = new PriemEntities())
                    {

                        Guid? entId = (from ent in MainClass.GetEntry(context)
                                       where ent.IsSecond == IsSecond && ent.IsParallel == IsParallel && ent.IsReduced == IsReduced
                                       && ent.LicenseProgramId == LicenseProgramId
                                       && ent.ObrazProgramId == ObrazProgramId
                                       && (ProfileId == null ? ent.ProfileId == 0 : ent.ProfileId == ProfileId)
                                       && ent.StudyFormId == StudyFormId
                                       && ent.StudyBasisId == StudyBasisId
                                       && ent.IsForeign == IsForeign
                                       select ent.Id).FirstOrDefault();
                        return entId;
                    }
                }
                catch
                {
                    return null;
                }
            }
        }
        private bool IsForeign
        {
            get { return MainClass.dbType == PriemType.PriemForeigners; }
        }
    }
}
