using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

using BDClassLib;
using EducServLib;

namespace Priem
{
    public class LoadFromInet
    {
        private DBPriem _bdcInet;

        public LoadFromInet()
        {
            _bdcInet = new DBPriem();
            try
            {
                _bdcInet.OpenDatabase(MainClass.connStringOnline);
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);                
            }
        }

        public DBPriem BDCInet
        {
            get { return _bdcInet; }
        }

        public void CloseDB()
        {
            _bdcInet.CloseDataBase();
        }          
        
        public void UpdatePersonData(int personBarc)
        {
            try
            {
                _bdcInet.ExecuteQuery("UPDATE Person SET IsImported = 1 WHERE Person.Barcode = " + personBarc);
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
        }

        public DataTable GetPersonEgeByBarcode(int fileNum)
        {
            string queryEge = "SELECT EgeMark.Id, EgeMark.EgeExamId AS ExamId, EgeMark.Value, EgeCertificate.Number, EgeMark.EgeCertificateId FROM EgeMark LEFT JOIN EgeCertificate ON EgeMark.EgeCertificateId = EgeCertificate.Id LEFT JOIN Person ON EgeCertificate.PersonId = Person.Id";
            DataSet dsEge = _bdcInet.GetDataSet(queryEge + " WHERE Person.Barcode = " + fileNum + " ORDER BY EgeMark.EgeCertificateId ");
            return dsEge.Tables[0];
        }

        public extPerson GetPersonByBarcode(int fileNum)
        {
            try
            {
                string personQueryInet =
                @"SELECT Id, Barcode, Name, SecondName, Surname, BirthDate, BirthPlace, Sex,
                    PassportTypeId, PassportSeries, PassportNumber, PassportAuthor, PassportDate,
                    PassportCode, '' AS PersonalCode, CountryId, NationalityId, RegionId, Phone, Mobiles, Email,
                    Code, City, Street, House, Korpus, Flat, CodeReal, CityReal, StreetReal, HouseReal, KorpusReal, FlatReal,
                    AbitHostel AS HostelAbit, LanguageId, 
                    Parents AS PersonInfo, AddInfo AS ExtraInfo, StartEnglish, EnglishMark, AbiturientTypeId, HostelEduc, SNILS, KladrCode
                    , HasTRKI, TRKICertificateNumber
                    FROM extPerson
                    WHERE 0=0";

                DataSet ds = _bdcInet.GetDataSet(personQueryInet + " AND extPerson.Barcode = " + fileNum);
                if (ds.Tables[0].Rows.Count == 0)
                    throw new Exception("Записей не найдено");

                DataRow row = ds.Tables[0].Rows[0];
                
                extPerson pers = new extPerson();
               
                int iAbitTypeId = (int)row["AbiturientTypeId"];

                //if (iAbitTypeId == 1 && MainClass.dbType == PriemType.PriemMag)
                //    WinFormsServ.Error("Выбран профиль человека, поступающего на 1 курс бакалавриата/специалитета");
                //if (iAbitTypeId == 2 && MainClass.dbType == PriemType.Priem)
                //    WinFormsServ.Error("Выбран профиль человека, поступающего в магистратуру");

                if (iAbitTypeId == 3 || iAbitTypeId == 4)
                    WinFormsServ.Error("Выбран профиль человека, переводящегося в СПбГУ. Проверьте штрих-код");
                if (iAbitTypeId == 5)
                    WinFormsServ.Error("Выбран профиль человека, восстанавливающегося в СПбГУ. Проверьте штрих-код");
                if (iAbitTypeId == 6 || iAbitTypeId == 7)
                    WinFormsServ.Error("Выбран профиль человека, переводящегося внутри СПбГУ. Проверьте штрих-код");
                if (iAbitTypeId == 8)
                    WinFormsServ.Error("Выбран профиль человека, поступающего в АГ СПбГУ. Проверьте штрих-код");
                if (iAbitTypeId == 9)
                    WinFormsServ.Error("Выбран профиль человека, поступающего в колледжи СПбГУ. Проверьте штрих-код");

                pers.Id = (Guid)row["Id"];
                pers.Barcode = (int?)row["Barcode"];
                pers.FIO = Util.GetFIO(row["Surname"].ToString(), row["Name"].ToString(), row["SecondName"].ToString());
                pers.Name = row["Name"].ToString();
                pers.SecondName = row["SecondName"].ToString();
                pers.Surname = row["Surname"].ToString();
                pers.BirthDate = QueryServ.ToNullDateTimeValue(row["BirthDate"]) ?? DateTime.Now;
                pers.BirthPlace = row["BirthPlace"].ToString();
                pers.PassportTypeId = (int?)(Util.ToNullObject(row["PassportTypeId"])) ?? 1;
                pers.PassportSeries = row["PassportSeries"].ToString();
                pers.PassportNumber = row["PassportNumber"].ToString();
                pers.PassportAuthor = row["PassportAuthor"].ToString();
                pers.PassportDate = QueryServ.ToNullDateTimeValue(row["PassportDate"]);
                pers.PassportCode = row["PassportCode"].ToString();
                pers.PersonalCode = row["PersonalCode"].ToString();
                pers.SNILS = row["SNILS"].ToString();
                
                pers.Sex = QueryServ.ToBoolValue(row["Sex"]);
                pers.CountryId = (int?)(Util.ToNullObject(row["CountryId"]));
                pers.NationalityId = (int?)(Util.ToNullObject(row["NationalityId"])) ?? 1;
                pers.RegionId = (int?)(Util.ToNullObject(row["RegionId"])) ?? 1;
                pers.Phone = row["Phone"].ToString();
                pers.Mobiles = row["Mobiles"].ToString();
                pers.Email = row["Email"].ToString();
                pers.Code = row["Code"].ToString();
                pers.City = row["City"].ToString();
                pers.Street = row["Street"].ToString();
                pers.House = row["House"].ToString();
                pers.Korpus = row["Korpus"].ToString();
                pers.Flat = row["Flat"].ToString();
                pers.CodeReal = row["CodeReal"].ToString();
                pers.CityReal = row["CityReal"].ToString();
                pers.StreetReal = row["StreetReal"].ToString();
                pers.HouseReal = row["HouseReal"].ToString();
                pers.KorpusReal = row["KorpusReal"].ToString();
                pers.FlatReal = row["FlatReal"].ToString();
                pers.KladrCode = pers.CountryId != 1 ? "иностр" : row["KladrCode"].ToString();
                pers.HostelAbit = QueryServ.ToBoolValue(row["HostelAbit"]);
                pers.HostelEduc = QueryServ.ToBoolValue(row["HostelEduc"]);
                pers.HasAssignToHostel = false;
                pers.HasExamPass = false;
                pers.LanguageId = (int?)(Util.ToNullObject(row["LanguageId"]));
                
                pers.HasTRKI = (bool)row["HasTRKI"];
                pers.TRKICertificateNumber = row["TRKICertificateNumber"].ToString();
                
                pers.PersonInfo = row["PersonInfo"].ToString();
                pers.ExtraInfo = row["ExtraInfo"].ToString();
                pers.StartEnglish = QueryServ.ToBoolValue(row["StartEnglish"]);
                int EnglishMark = 0;
                int.TryParse(row["EnglishMark"].ToString(), out EnglishMark);
                pers.EnglishMark = EnglishMark == 0 ? null : (int?)EnglishMark;

                DataSet dsWork = _bdcInet.GetDataSet(string.Format(@"
                      SELECT  PersonWork.WorkPlace + ', ' + PersonWork.WorkProfession + ', ' + PersonWork.WorkSpecifications + ' стаж: ' + PersonWork.Stage AS Work,
                      PersonWork.WorkPlace + ', ' + PersonWork.WorkProfession + ', ' + PersonWork.WorkSpecifications AS Place, PersonWork.Stage
                      FROM PersonWork WHERE PersonWork.PersonId = '{0}'", pers.Id));

                if (dsWork.Tables[0].Rows.Count == 0)
                {
                    pers.Stag = string.Empty;
                    pers.WorkPlace = string.Empty;
                }
                else if (dsWork.Tables[0].Rows.Count == 1)
                {
                    pers.Stag = dsWork.Tables[0].Rows[0]["Stage"].ToString();
                    pers.WorkPlace = dsWork.Tables[0].Rows[0]["Place"].ToString();
                }
                else
                {
                    string work = string.Empty;
                    foreach (DataRow dr in dsWork.Tables[0].Rows)
                    {
                        work += dr["Work"].ToString() + ";" + Environment.NewLine;
                    }
                    pers.WorkPlace = work;
                }

                DataSet dsScienceWork = _bdcInet.GetDataSet(string.Format(@"
                      SELECT ScienceWorkType.Name + ': ' + PersonScienceWork.WorkInfo AS ScienseWork                      
                      FROM PersonScienceWork LEFT JOIN ScienceWorkType ON PersonScienceWork.WorkTypeId = ScienceWorkType.Id WHERE PersonScienceWork.PersonId = '{0}'", pers.Id));

                if (dsScienceWork.Tables[0].Rows.Count == 0)
                    pers.ScienceWork = string.Empty;                
               
                else
                {
                    string work = string.Empty;
                    foreach (DataRow dr in dsScienceWork.Tables[0].Rows)
                    {
                        work += dr["ScienseWork"].ToString() + ";" + Environment.NewLine;
                    }
                    pers.ScienceWork = work;
                }
               
                return pers;
            }
            catch
            {
                return null;
            }
        }

        public List<Person_EducationInfo> GetPersonEducationDocumentsByBarcode(int fileNum)
        {
            List<Person_EducationInfo> lstRet = new List<Person_EducationInfo>();

            string query = @"select
Person.Id as PersonId, PersonEducationDocument.Id as EducationDocumentId
, SchoolCity, SchoolTypeId, SchoolName, SchoolNum, SchoolExitYear
, Country.PriemDictionaryId as CountryEducId, Region.PriemDictionaryId AS RegionEducId
, IsEqual, EqualDocumentNumber, Series , Number, AvgMark, IsExcellent
, PersonHighEducationInfo.EducationDocumentId as PersonHighEducationInfoId
, Qualification.Name as Qualification, EntryYear, ExitYear, DiplomaTheme, QualificationId, StudyFormId, ProgramName
from dbo.PersonEducationDocument 
inner join Person on Person.Id = PersonEducationDocument.PersonId
inner join Country on Country.Id = PersonEducationDocument.CountryEducId
inner join Region on Region.Id = PersonEducationDocument.RegionEducId
left join PersonHighEducationInfo on PersonHighEducationInfo.EducationDocumentId = PersonEducationDocument.Id
left join Qualification on Qualification.Id = QualificationId
where Person.Barcode =" + fileNum ;
            
            DataSet ds_EducationInfo = _bdcInet.GetDataSet(query);

            if (ds_EducationInfo.Tables[0].Rows.Count == 0)
                throw new Exception("Записей не найдено");

            foreach (DataRow row in ds_EducationInfo.Tables[0].Rows)
            {
                lstRet.Add(
                    new Person_EducationInfo()
                    {
                        Id = row.Field<int>("EducationDocumentId"),
                        PersonId = row.Field<Guid>("PersonId"),
                        SchoolCity = row.Field<string>("SchoolCity"),
                        SchoolTypeId = row.Field<int>("SchoolTypeId"),
                        SchoolName = row.Field<string>("SchoolName"),
                        SchoolNum = row.Field<string>("SchoolNum"),
                        SchoolExitYear = int.Parse(row.Field<string>("SchoolExitYear")),
                        CountryEducId = row.Field<int>("CountryEducId"),
                        RegionEducId = row.Field<int>("RegionEducId"),
                        IsExcellent = row.Field<bool>("IsExcellent"),
                        IsEqual = row.Field<bool>("IsEqual"),
                        EqualDocumentNumber = row.Field<string>("EqualDocumentNumber"),
                        AttestatSeries = row.Field<string>("Series"),
                        AttestatNum = row.Field<string>("Number"),
                        DiplomSeries = row.Field<string>("Series"),
                        DiplomNum = row.Field<string>("Number"),
                        SchoolAVG = row.Field<double>("AvgMark"),
                        HighEducation = row.Field<string>("SchoolName"),
                        HEProfession = row.Field<string>("ProgramName") ?? "",
                        HEQualification = row.Field<string>("Qualification") ?? " ",
                        HEEntryYear = row.Field<int?>("EntryYear"),
                        HEExitYear = row.Field<int?>("ExitYear"),
                        HEWork = row.Field<string>("DiplomaTheme") ?? "",
                        HEStudyFormId = row.Field<int?>("StudyFormId"),
                    });
            }

            return lstRet;
        }

        public List<ShortCompetition> GetCompetitionList(int _abitBarc)
        {
            List<ShortCompetition> LstCompetitions = new List<ShortCompetition>();
            try
            {
                string query =
    @"SELECT Abiturient.[Id]
,[Priority]
,[PersonId]
,[Priority]
,[Barcode]
,[DateOfStart]
,[EntryId]
,[FacultyId]
,[FacultyName]
,[LicenseProgramId]
,[LicenseProgramCode]
,[LicenseProgramName]
,[ObrazProgramId]
,[ObrazProgramCrypt]
,[ObrazProgramName]
,[ProfileId]
,[ProfileName]
,[StudyBasisId]
,[StudyBasisName]
,[StudyFormId]
,[StudyFormName]
,[StudyLevelId]
,[StudyLevelName]
,[IsSecond]
,[IsReduced]
,[IsParallel]
,[IsGosLine]
,[CommitId]
,[DateOfStart]
,(SELECT MAX(ApplicationCommitVersion.Id) FROM ApplicationCommitVersion WHERE ApplicationCommitVersion.CommitId = [Abiturient].CommitId) AS VersionNum
,(SELECT MAX(ApplicationCommitVersion.VersionDate) FROM ApplicationCommitVersion WHERE ApplicationCommitVersion.CommitId = [Abiturient].CommitId) AS VersionDate
,ApplicationCommit.IntNumber
,[Abiturient].HasInnerPriorities
,[Abiturient].IsApprovedByComission
,[Abiturient].CompetitionId
,[Abiturient].ApproverName
,[Abiturient].DocInsertDate
,[Abiturient].IsCommonRussianCompetition
FROM [Abiturient] 
INNER JOIN ApplicationCommit ON ApplicationCommit.Id = Abiturient.CommitId
WHERE IsCommited = 1 AND IntNumber=@CommitId";

                DataTable tbl = _bdcInet.GetDataSet(query, new SortedList<string, object>() { { "@CommitId", _abitBarc } }).Tables[0];

                LstCompetitions =
                             (from DataRow rw in tbl.Rows
                              select new ShortCompetition(rw.Field<Guid>("Id"), rw.Field<Guid>("CommitId"), rw.Field<Guid>("EntryId"), rw.Field<Guid>("PersonId"),
                                  rw.Field<int?>("VersionNum"), rw.Field<DateTime?>("VersionDate"))
                              {
                                  Barcode = rw.Field<int>("Barcode"),
                                  CompetitionId = rw.Field<int?>("CompetitionId") ?? (rw.Field<int>("StudyBasisId") == 1 ? 4 : 3),
                                  CompetitionName = "не указана",
                                  HasCompetition = rw.Field<bool>("IsApprovedByComission"),
                                  LicenseProgramId = rw.Field<int>("LicenseProgramId"),
                                  LicenseProgramName = rw.Field<string>("LicenseProgramName"),
                                  ObrazProgramId = rw.Field<int>("ObrazProgramId"),
                                  ObrazProgramName = rw.Field<string>("ObrazProgramName"),
                                  ProfileId = rw.Field<int?>("ProfileId") ?? 0,
                                  ProfileName = rw.Field<string>("ProfileName"),
                                  StudyBasisId = rw.Field<int>("StudyBasisId"),
                                  StudyBasisName = rw.Field<string>("StudyBasisName"),
                                  StudyFormId = rw.Field<int>("StudyFormId"),
                                  StudyFormName = rw.Field<string>("StudyFormName"),
                                  StudyLevelId = rw.Field<int>("StudyLevelId"),
                                  StudyLevelName = rw.Field<string>("StudyLevelName"),
                                  FacultyId = rw.Field<int>("FacultyId"),
                                  FacultyName = rw.Field<string>("FacultyName"),
                                  DocDate = rw.Field<DateTime>("DateOfStart"),
                                  DocInsertDate = rw.Field<DateTime?>("DocInsertDate") ?? DateTime.Now,
                                  Priority = rw.Field<int>("Priority"),
                                  IsGosLine = rw.Field<bool>("IsGosLine"),
                                  IsReduced = rw.Field<bool>("IsReduced"),
                                  IsSecond = rw.Field<bool>("IsSecond"),
                                  HasInnerPriorities = rw.Field<bool>("HasInnerPriorities"),
                                  IsApprovedByComission = rw.Field<bool>("IsApprovedByComission"),
                                  ApproverName = rw.Field<string>("ApproverName"),
                                  lstInnerEntryInEntry = new List<ShortInnerEntryInEntry>(),
                                  IsCommonRussianCompetition = rw.Field<bool>("IsCommonRussianCompetition"),
                              }).ToList();

                //ObrazProgramInEntry
                foreach (var C in LstCompetitions.Where(x => x.HasInnerPriorities))
                {
                    C.lstInnerEntryInEntry = new List<ShortInnerEntryInEntry>();
                    query = @"SELECT InnerEntryInEntryId, InnerEntryInEntryPriority, ObrazProgramName, ProfileName, 
ISNULL(CurrVersion, 1) AS CurrVersion, ISNULL(CurrDate, GETDATE()) AS CurrDate
FROM [extApplicationDetails] WHERE [ApplicationId]=@AppId";
                    tbl = _bdcInet.GetDataSet(query, new SortedList<string, object>() { { "@AppId", C.Id } }).Tables[0];

                    var data = (from DataRow rw in tbl.Rows
                                select new
                                {
                                    InnerEntryInEntryId = rw.Field<Guid>("InnerEntryInEntryId"),
                                    InnerEntryInEntryPriority = rw.Field<int>("InnerEntryInEntryPriority"),
                                    ObrazProgramName = rw.Field<string>("ObrazProgramName"),
                                    ProfileName = rw.Field<string>("ProfileName"),
                                    CurrVersion = rw.Field<int>("CurrVersion"),
                                    CurrDate = rw.Field<DateTime>("CurrDate")
                                }).ToList().OrderBy(x => x.InnerEntryInEntryPriority).ToList();

                    foreach (var OPIE in data)
                    {
                        var OP = new ShortInnerEntryInEntry(OPIE.InnerEntryInEntryId, OPIE.ObrazProgramName, OPIE.ProfileName);
                        OP.InnerEntryInEntryPriority = OPIE.InnerEntryInEntryPriority;
                        OP.CurrVersion = OPIE.CurrVersion;
                        OP.CurrDate = OPIE.CurrDate;
                        C.lstInnerEntryInEntry.Add(OP);
                    }
                }
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при заполнении формы заявления" + ex.Message);
            }

            return LstCompetitions;
        }

        public void UpdateApplicationSetApprovedByComission(ShortCompetition comp)
        {
            string query = @"UPDATE [Application] SET IsApprovedByComission=1, ApproverName=@ApproverName, CompetitionId=@CompId, DocInsertDate=@DocInsertDate, 
IsCommonRussianCompetition=@IsCommonRussianCompetition, IsGosLine=@IsGosLine WHERE Id=@Id";
            _bdcInet.ExecuteQuery(query, new SortedList<string, object>()
                {
                    { "@Id", comp.Id },
                    { "@CompId", comp.CompetitionId },
                    { "@DocInsertDate", comp.DocInsertDate },
                    { "@ApproverName", MainClass.GetUserName() },
                    { "@IsGosLine", comp.IsGosLine },
                    { "@IsCommonRussianCompetition", comp.IsCommonRussianCompetition }
                });
        }

        public void UpdateApplicationCommitSetIsImported(int? _abitBarc)
        {
            if (!MainClass.IsTestDB)
                _bdcInet.ExecuteQuery("UPDATE ApplicationCommit SET IsImported = 1 WHERE IntNumber = '" + _abitBarc + "'");
        }
    }
}
