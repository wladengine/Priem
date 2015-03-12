using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Priem
{
    public static class ProtocolDataProvider
    {
        public static List<AbitDataInProtocol> GetProtocolData(Guid ProtocolId)
        {
            using (PriemEntities ctx = new PriemEntities())
            {
                List<AbitDataInProtocol> lst =
                    (from extabit in ctx.extAbit
                     join extperson in ctx.extPerson on extabit.PersonId equals extperson.Id
                     join currEduc in ctx.extPerson_EducationInfo_Current on extperson.Id equals currEduc.PersonId into currEduc2
                     from currEduc in currEduc2.DefaultIfEmpty()
                     join qentry in ctx.qEntry on extabit.EntryId equals qentry.Id
                     join competition in ctx.Competition on extabit.CompetitionId equals competition.Id into competition2
                     from competition in competition2.DefaultIfEmpty()
                     join extprotocol in ctx.extProtocol on extabit.Id equals extprotocol.AbiturientId into extprotocol2
                     from extprotocol in extprotocol2.DefaultIfEmpty()
                     where extprotocol.Id == ProtocolId
                     orderby qentry.LicenseProgramCode + " " + qentry.LicenseProgramName + ", " + qentry.ObrazProgramName + ", " + (qentry.ProfileId != null ? qentry.ProfileName : ""), extabit.RegNum
                     select new
                     {
                         AbiturientId = extabit.Id,
                         RegNum = extabit.RegNum,
                         FIO = extperson.Surname + " " + extperson.Name + " " + extperson.SecondName,
                         EducationDocument = currEduc.SchoolTypeId == 1 ? currEduc.AttestatSeries + "  №" + currEduc.AttestatNum : currEduc.DiplomSeries + "  №" + currEduc.DiplomNum,
                         Direction = qentry.LicenseProgramCode + " " + qentry.LicenseProgramName + ", " + qentry.ObrazProgramName + ", " + (qentry.ProfileId != null ? qentry.ProfileName : ""),
                         LicenseProgramCode = qentry.LicenseProgramCode,
                         CompetitionName = competition.Name,
                         PersonId = extabit.PersonId,
                         EntryId = extabit.EntryId,
                         Comment = extabit.BackDoc ? "Забрал док." : extabit.NotEnabled ? "Не допущен" : extprotocol.Excluded ? "Исключён из протокола" : ""
                     }).ToList().Distinct()
                     .Select(x => new AbitDataInProtocol
                     {
                         AbiturientId = x.AbiturientId,
                         PersonId = x.PersonId,
                         EntryId = x.EntryId,
                         RegNum = x.RegNum,
                         FIO = x.FIO,
                         EducationDocument = x.EducationDocument,
                         Direction = x.Direction,
                         LicenseProgramCode = x.LicenseProgramCode,
                         CompetitionName = x.CompetitionName,
                         Comment = x.Comment
                     }).ToList();

                return lst;
            }
        }
        public static List<AbitDataInEntryView> GetEntryViewData(Guid ProtocolId, bool? isRus)
        {
            using (PriemEntities ctx = new PriemEntities())
            {
                List<AbitDataInEntryView> lst =
                    (from extabit in ctx.extAbit
                     join extentryView in ctx.extEntryView on extabit.Id equals extentryView.AbiturientId
                     join extperson in ctx.extPerson on extabit.PersonId equals extperson.Id
                     join country in ctx.Country on extperson.NationalityId equals country.Id
                     join competition in ctx.Competition on extabit.CompetitionId equals competition.Id
                     join extabitMarksSum in ctx.extAbitMarksSum on extabit.Id equals extabitMarksSum.Id into extabitMarksSum2
                     from extabitMarksSum in extabitMarksSum2.DefaultIfEmpty()
                     join entryHeader in ctx.EntryHeader on extentryView.EntryHeaderId equals entryHeader.Id into entryHeader2
                     from entryHeader in entryHeader2.DefaultIfEmpty()
                     join celCompetition in ctx.CelCompetition on extabit.CelCompetitionId equals celCompetition.Id into celCompetition2
                     from celCompetition in celCompetition2.DefaultIfEmpty()
                     where extentryView.Id == ProtocolId && (isRus.HasValue ? (isRus.Value ? extperson.NationalityId == 1 : extperson.NationalityId != 1) : true)
                     orderby celCompetition.TvorName, extabit.ObrazProgramName, extabit.ProfileName, country.NameRod, entryHeader.SortNum, extabit.FIO
                     select new
                     {
                         AbiturientId = extabit.Id,
                         extabit.RegNum,
                         extabit.PersonNum,
                         TotalSum = (extabit.CompetitionId == 8 || extabit.CompetitionId == 1) ? null : extabitMarksSum.TotalSum,
                         extabit.FIO,
                         CelCompName = celCompetition.TvorName,
                         LicenseProgramName = extabit.LicenseProgramName,
                         LicenseProgramCode = extabit.LicenseProgramCode,
                         ProfileName = extabit.ProfileName,
                         ObrazProgram = extabit.ObrazProgramName,
                         ObrazProgramId = extabit.ObrazProgramId,
                         EntryHeaderId = entryHeader.Id,
                         SortNum = entryHeader.SortNum,
                         EntryHeaderName = entryHeader.Name,
                         CountryNameRod = country.NameRod,
                         extabit.InnerEntryInEntryObrazProgramId,
                         extabit.InnerEntryInEntryObrazProgramCrypt,
                         extabit.InnerEntryInEntryObrazProgramName,
                         InnerEntryInEntryProfileName = extabit.InnerEntryInEntryProfileName,
                         extabit.CompetitionId,
                         extentryView.SignerName,
                         extentryView.SignerPosition,
                         extabit.ObrazProgramCrypt
                     }).ToList().Distinct()
                     .Select(x => new AbitDataInEntryView
                     {
                         AbiturientId = x.AbiturientId,
                         RegNum = x.RegNum,
                         PersonNum = x.PersonNum,
                         TotalSum = x.TotalSum,
                         FIO = x.FIO,
                         CelCompName = x.CelCompName,
                         LicenseProgramName = x.LicenseProgramName,
                         LicenseProgramCode = x.LicenseProgramCode,
                         ProfileName = string.IsNullOrEmpty(x.InnerEntryInEntryProfileName) ? x.ProfileName : x.InnerEntryInEntryProfileName,
                         ObrazProgram = x.InnerEntryInEntryObrazProgramId.HasValue ? x.InnerEntryInEntryObrazProgramCrypt + " " + x.InnerEntryInEntryObrazProgramName : x.ObrazProgram.Replace("(очно-заочная)", "").Replace(" ВВ", ""),
                         ObrazProgramId = x.InnerEntryInEntryObrazProgramId.HasValue ? x.InnerEntryInEntryObrazProgramId.Value : x.ObrazProgramId,
                         EntryHeaderId = x.EntryHeaderId,
                         EntryHeaderName = x.EntryHeaderName,
                         SortNum = x.SortNum,
                         CountryNameRod = x.CountryNameRod,
                         InnerEntryInEntryObrazProgramCrypt = x.InnerEntryInEntryObrazProgramCrypt,
                         InnerEntryInEntryObrazProgramName = x.InnerEntryInEntryObrazProgramName,
                         InnerEntryInEntryProfileName = x.InnerEntryInEntryProfileName,
                         CompetitionId = x.CompetitionId,
                         SignerName = x.SignerName,
                         SignerPosition = x.SignerPosition,
                         ObrazProgramCrypt = x.InnerEntryInEntryObrazProgramId.HasValue ? x.InnerEntryInEntryObrazProgramCrypt : x.ObrazProgramCrypt
                    }).OrderBy(x => x.CelCompName).ThenBy(x => x.ObrazProgram).ThenBy(x => x.ProfileName).ThenBy(x => x.CountryNameRod).ThenBy(x => x.SortNum).ThenBy(x => x.FIO)
                    .ToList();

                return lst;
            }
        }

        public static ProtocolInfo GetProtocolInfo(Guid ProtocolId, int ProtocolTypeId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                var info =
                    (from protocol in context.Protocol
                     join sf in context.StudyForm on protocol.StudyFormId equals sf.Id
                     where protocol.Id == ProtocolId
                     && protocol.ProtocolTypeId == ProtocolTypeId && protocol.IsOld == false
                     select new ProtocolInfo
                     {
                         StudyFormName = sf.Name,
                         StudyFormId = protocol.StudyFormId,
                         StudyBasisId = protocol.StudyBasisId,
                         Date = protocol.Date,
                         Number = protocol.Number,
                         IsListener = protocol.IsListener,
                         IsReduced = protocol.IsReduced,
                         IsSecond = protocol.IsSecond,
                         IsParallel = protocol.IsParallel,
                         FacultyName = protocol.FacultyId.HasValue ? protocol.SP_Faculty.Name : "",
                         FacultyDatName = protocol.FacultyId.HasValue ? protocol.SP_Faculty.DatName : "",
                         LicenseProgramCode = protocol.LicenseProgramId.HasValue ? protocol.SP_LicenseProgram.Code : "",
                         LicenseProgramName = protocol.LicenseProgramId.HasValue ? protocol.SP_LicenseProgram.Name : "",
                         StudyLevelId = protocol.LicenseProgramId.HasValue ? protocol.SP_LicenseProgram.StudyLevelId : 0,
                         StudyLevelName = protocol.LicenseProgramId.HasValue ? protocol.SP_LicenseProgram.StudyLevel.Name : "",
                         StudyLevelNameRod = protocol.LicenseProgramId.HasValue ? protocol.SP_LicenseProgram.StudyLevel.NameRod : "",
                     }).FirstOrDefault();

                return info;
            }
        }

        public class AbitDataInProtocol
        {
            public Guid AbiturientId { get; set; }
            public Guid PersonId { get; set; }
            public Guid EntryId { get; set; }
            public string RegNum { get; set; }
            public string FIO { get; set; }
            public string EducationDocument { get; set; }
            public string Direction { get; set; }
            public string LicenseProgramCode { get; set; }
            public string CompetitionName { get; set; }
            public string Comment { get; set; }
        }
        public class AbitDataInEntryView
        {
            public Guid AbiturientId { get; set; }
            public string RegNum { get; set; }
            public string PersonNum { get; set; }
            public string FIO { get; set; }
            public int? TotalSum { get; set; }

            public string LicenseProgramName { get; set; }
            public string LicenseProgramCode { get; set; }

            public int ObrazProgramId { get; set; }
            public string ObrazProgram { get; set; }
            public string ObrazProgramCrypt { get; set; }
            public string ProfileName { get; set; }

            public int EntryHeaderId { get; set; }
            public string EntryHeaderName { get; set; }

            public int SortNum { get; set; }
            public string CountryNameRod { get; set; }
            public string CelCompName { get; set; }

            public int CompetitionId { get; set; }

            public string InnerEntryInEntryObrazProgramCrypt { get; set; }
            public string InnerEntryInEntryObrazProgramName { get; set; }
            public string InnerEntryInEntryProfileName { get; set; }

            public string SignerName { get; set; }
            public string SignerPosition { get; set; }
        }
        public class ProtocolInfo
        {
            public string Number { get; set; }
            public DateTime Date { get; set; }
            public int? StudyBasisId { get; set; }
            
            public int StudyLevelId { get; set; }
            public string StudyLevelName { get; set; }
            public string StudyLevelNameRod { get; set; }

            public int StudyFormId { get; set; }
            public string StudyFormName { get; set; }

            public string FacultyName { get; set; }
            public string FacultyDatName { get; set; }

            public string LicenseProgramName { get; set; }
            public string LicenseProgramCode { get; set; }

            public bool IsListener { get; set; }
            public bool IsSecond { get; set; }
            public bool IsReduced { get; set; }
            public bool IsParallel { get; set; }
        }
    }
}
