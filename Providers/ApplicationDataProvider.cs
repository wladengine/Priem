using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Priem
{
    static class ApplicationDataProvider
    {
        public static List<ApplicationData> GetAppData(Guid PersonId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<ApplicationData> abitList =
                    (from x in context.Abiturient
                     join Ent in context.Entry on x.EntryId equals Ent.Id
                     where Ent.StudyLevel.StudyLevelGroup.Id == MainClass.studyLevelGroupId
                     && x.IsGosLine == false
                     && x.PersonId == PersonId
                     && x.BackDoc == false
                     select new ApplicationData
                     {
                         ApplicationId = x.Id,
                         PersonId = x.PersonId,
                         Barcode = x.Barcode,
                         FacultyName = Ent.SP_Faculty.Name,
                         LicenseProgramName = Ent.SP_LicenseProgram.Name,
                         LicenseProgramCode = Ent.SP_LicenseProgram.Code,
                         ObrazProgramCrypt = Ent.StudyLevel.Acronym + "." + Ent.SP_ObrazProgram.Number + "." + MainClass.sPriemYear,
                         ObrazProgramName = Ent.SP_ObrazProgram.Name,
                         ProfileName = Ent.SP_Profile.Name,
                         StudyFormId = Ent.StudyFormId,
                         StudyFormName = Ent.StudyForm.Name,
                         StudyBasisId = Ent.StudyBasisId,
                         StudyLevelId = Ent.StudyLevelId,
                         Priority = x.Priority,
                         IsGosLine = x.IsGosLine,
                         ComissionId = Ent.CommissionId,
                         ComissionAddress = Ent.CommissionId.HasValue ? Ent.Comission.Name : ""
                     }).OrderBy(x => x.Priority).ToList();

                return abitList;
            }
        }
        public static List<ShortAppcationDetails> GetAppProfileList(Guid PersonId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<ShortAppcationDetails> abitProfileList =
                    (from x in context.Abiturient
                     join Ad in context.extApplicationDetails on x.Id equals Ad.ApplicationId
                     join Ent in context.Entry on x.EntryId equals Ent.Id
                     where Ent.StudyLevel.StudyLevelGroup.Id == MainClass.studyLevelGroupId
                     && x.IsGosLine == false
                     && x.PersonId == PersonId
                     && x.BackDoc == false
                     select new ShortAppcationDetails()
                     {
                         ApplicationId = x.Id,
                         Priority = Ad.InnerEntryInEntryPriority,
                         ObrazProgramName = Ad.ObrazProgramCrypt + " " + Ad.ObrazProgramName,
                         ProfileName = Ad.ProfileName
                     }).Distinct().ToList();

                return abitProfileList;
            }
        }
    }

    public class ApplicationData
    {
        public Guid ApplicationId { get; set; }
        public Guid PersonId { get; set; }
        public Guid EntryId { get; set; }
        public int? Barcode { get; set; }
        public int? Priority { get; set; }

        public bool IsGosLine { get; set; }

        public int FacultyId { get; set; }
        public string FacultyName { get; set; }

        public int StudyLevelId { get; set; }
        public string StudyLevelName { get; set; }

        public int LicenseProgramId { get; set; }
        public string LicenseProgramName { get; set; }
        public string LicenseProgramCode { get; set; }

        public int ObrazProgramId { get; set; }
        public string ObrazProgramName { get; set; }
        public string ObrazProgramCrypt { get; set; }

        public int ProfileId { get; set; }
        public string ProfileName { get; set; }

        public int StudyFormId { get; set; }
        public string StudyFormName { get; set; }

        public int StudyBasisId { get; set; }
        public string StudyBasisName { get; set; }

        public int? ComissionId { get; set; }
        public string ComissionAddress { get; set; }
    }
}
