using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Priem
{
    public static class ApplicationCommitSaveProvider
    {
        public static void CheckAndUpdateNotUsedApplications(Guid personId, List<ShortCompetition> LstCompetitions)
        {
            using (PriemEntities context = new PriemEntities())
            {
                var notUsedApplications = context.Abiturient
                    .Where(x => x.PersonId == personId && !x.BackDoc && x.Entry.StudyLevel.LevelGroupId == MainClass.studyLevelGroupId)
                    .Select(x => x.EntryId).ToList()
                    .Except(LstCompetitions.Select(x => x.EntryId)).ToList();
                if (notUsedApplications.Count > 0)
                {
                    var dr = MessageBox.Show("У абитуриента в базе имеются " + notUsedApplications.Count +
                        " конкурсов, не перечисленных в заявлении. Вероятно, по ним был уже произведён отказ. Проставить по данным конкурсным позициям отказ от участия в конкурсе?",
                        "Внимание!", MessageBoxButtons.YesNo);
                    if (dr == System.Windows.Forms.DialogResult.Yes)
                    {
                        string str = "У меня есть на руках заявление об отказе в участии в следующих конкурсах:";
                        int incrmntr = 1;
                        foreach (var app_entry in notUsedApplications)
                        {
                            var entry = context.Entry.Where(x => x.Id == app_entry).FirstOrDefault();
                            str += "\n" + incrmntr++ + ")" + entry.SP_LicenseProgram.Code + " " + entry.SP_LicenseProgram.Name + "; "
                                + entry.StudyLevel.Acronym + "." + entry.SP_ObrazProgram.Number + " " + entry.SP_ObrazProgram.Name +
                                ";\nПрофиль:" + entry.SP_Profile.Name + ";" + entry.StudyForm.Acronym + ";" + entry.StudyBasis.Acronym;
                        }
                        
                        dr = MessageBox.Show(str, "Внимание!", MessageBoxButtons.YesNo);
                        if (dr == System.Windows.Forms.DialogResult.Yes)
                        {
                            foreach (var app_entry in notUsedApplications)
                            {
                                var applst = context.Abiturient.Where(x => x.EntryId == app_entry && x.PersonId == personId && !x.BackDoc && x.Entry.StudyLevel.LevelGroupId == MainClass.studyLevelGroupId).Select(x => x.Id).ToList();
                                foreach (var app in applst)
                                    context.Abiturient_UpdateBackDoc(true, DateTime.Now, app);
                            }
                        }
                    }
                }
            }
        }
        public static void SaveApplicationCommitInWorkBase(Guid PersonId, List<ShortCompetition> LstCompetitions, int? LanguageId, int? _abitBarc)
        {
            using (PriemEntities context = new PriemEntities())
            {
                foreach (var Comp in LstCompetitions)
                {
                    var DocDate = Comp.DocDate;
                    var DocInsertDate = Comp.DocInsertDate == DateTime.MinValue ? DateTime.Now : Comp.DocInsertDate;

                    bool isViewed = Comp.HasCompetition;
                    Guid ApplicationId = Comp.Id;
                    bool hasLoaded = context.Abiturient.Where(x => x.PersonId == PersonId && x.EntryId == Comp.EntryId && !x.BackDoc).Count() == 0;
                    if (hasLoaded)
                    {
                        context.Abiturient_InsertDirectly(PersonId, Comp.EntryId, Comp.CompetitionId, Comp.IsListener,
                            false, false, false, null, DocDate, DocInsertDate,
                            false, false, null, Comp.OtherCompetitionId, Comp.CelCompetitionId, Comp.CelCompetitionText,
                            LanguageId, Comp.HasOriginals, Comp.Priority, Comp.Barcode, Comp.CommitId, _abitBarc, Comp.IsGosLine, isViewed, ApplicationId);
                        context.Abiturient_UpdateIsCommonRussianCompetition(Comp.IsCommonRussianCompetition, ApplicationId);
                    }
                    else
                    {
                        ApplicationId = context.Abiturient.Where(x => x.PersonId == PersonId && x.EntryId == Comp.EntryId && !x.BackDoc).Select(x => x.Id).First();
                        context.Abiturient_UpdatePriority(Comp.Priority, ApplicationId);
                    }
                    if (Comp.lstInnerEntryInEntry.Count > 0)
                    {
                        //загружаем внутренние приоритеты по профилям
                        int currVersion = Comp.lstInnerEntryInEntry.Select(x => x.CurrVersion).FirstOrDefault();
                        DateTime currDate = Comp.lstInnerEntryInEntry.Select(x => x.CurrDate).FirstOrDefault();
                        Guid ApplicationVersionId = Guid.NewGuid();
                        context.ApplicationVersion.AddObject(new ApplicationVersion() { IntNumber = currVersion, Id = ApplicationVersionId, ApplicationId = ApplicationId, VersionDate = currDate });
                        foreach (var InnerEntryInEntry in Comp.lstInnerEntryInEntry)
                        {
                            context.Abiturient_UpdateInnerEntryInEntryPriority(InnerEntryInEntry.Id, InnerEntryInEntry.InnerEntryInEntryPriority, ApplicationId);
                        }

                        context.SaveChanges();
                    }
                }
            }
        }
    }
}
