using System;
using System.Collections.Generic;
using System.Data.Objects;
using System.Linq;
using System.Text;

namespace Priem
{
    static class PersonDataProvider
    {
        public static List<Person_EducationInfo> GetPersonEducationDocumentsById(Guid PersonId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<Person_EducationInfo> lstRet = new List<Person_EducationInfo>();
                var EducationInfo = (from pEd in context.Person_EducationInfo
                                     where pEd.PersonId == PersonId
                                     select pEd).ToList();

                if (EducationInfo.Count == 0)
                    throw new Exception("Не найдено записей о документах об образовании");

                foreach (var row in EducationInfo)
                {
                    lstRet.Add(
                        new Person_EducationInfo()
                        {
                            Id = row.Id,
                            PersonId = row.PersonId,
                            SchoolCity = row.SchoolCity,
                            SchoolTypeId = row.SchoolTypeId,
                            SchoolName = row.SchoolName,
                            SchoolNum = row.SchoolNum,
                            SchoolExitYear = row.SchoolExitYear,
                            CountryEducId = row.CountryEducId,
                            RegionEducId = row.RegionEducId,
                            IsExcellent = row.IsExcellent,
                            IsEqual = row.IsEqual,
                            EqualDocumentNumber = row.EqualDocumentNumber,
                            AttestatSeries = row.AttestatSeries,
                            AttestatNum = row.AttestatNum,
                            DiplomSeries = row.DiplomSeries,
                            DiplomNum = row.DiplomNum,
                            SchoolAVG = row.SchoolAVG,
                            HighEducation = row.HighEducation,
                            HEProfession = row.HEProfession,
                            HEQualification = row.HEQualification,
                            HEEntryYear = row.HEEntryYear,
                            HEExitYear = row.HEExitYear,
                            HEWork = row.HEWork,
                            HEStudyFormId = row.HEStudyFormId,
                        }
                        );
                }

                return lstRet;
            }
        }

        public static void DeletePersonEducationDocument(Guid PersonId, int id)
        {
            using (PriemEntities context = new PriemEntities())
            {
                context.Person_EducationInfo_delete(PersonId, id);
            }
        }

        public static int CreateNewEducationInfo(Guid PersonId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                int? CountryEducId = MainClass.countryRussiaId;
                int? RegionEducId = context.Region.Select(x => (int)x.Id).FirstOrDefault();

                var Person = context.Person_Contacts.Where(x => x.PersonId == PersonId).FirstOrDefault();
                if (Person != null)
                {
                    CountryEducId = Person.CountryId;
                    RegionEducId = Person.RegionId;
                }

                int? SchoolTypeId = context.SchoolType.Select(x => (int)x.Id).FirstOrDefault();

                ObjectParameter idParam = new ObjectParameter("id", typeof(int));
                context.Person_EducationInfo_insert(PersonId, false, "", SchoolTypeId, "", "", null, null,
                    CountryEducId, RegionEducId, false, "", "", "", "", "", "", "", null, null, null, "", idParam);

                return (int)idParam.Value;
            }
        }

        public static void SaveEducationDocument(Person_EducationInfo ED)
        {
            ObjectParameter idParam = new ObjectParameter("id", typeof(int));
            using (PriemEntities context = new PriemEntities())
            {
                if (ED.Id == 0)
                    context.Person_EducationInfo_insert(ED.PersonId, ED.IsExcellent, ED.SchoolCity, ED.SchoolTypeId, ED.SchoolName,
                        ED.SchoolNum, ED.SchoolExitYear, ED.SchoolAVG, ED.CountryEducId, ED.RegionEducId, ED.IsEqual,
                        ED.AttestatSeries, ED.AttestatNum, ED.DiplomSeries, ED.DiplomNum, ED.HighEducation,
                        ED.HEProfession, ED.HEQualification, ED.HEEntryYear, ED.HEExitYear, ED.HEStudyFormId, ED.HEWork, idParam);
                else
                    context.Person_EducationInfo_update(ED.PersonId, ED.IsExcellent, ED.SchoolCity, ED.SchoolTypeId, ED.SchoolName,
                        ED.SchoolNum, ED.SchoolExitYear, ED.SchoolAVG, ED.CountryEducId, ED.RegionEducId, ED.IsEqual,
                        ED.AttestatSeries, ED.AttestatNum, ED.DiplomSeries, ED.DiplomNum, ED.HighEducation,
                        ED.HEProfession, ED.HEQualification, ED.HEEntryYear, ED.HEExitYear, ED.HEStudyFormId, ED.HEWork, ED.Id);
            }
        }
        
        /// <summary>
        /// Возвращает, есть ли человек в представлении к зачислению
        /// </summary>
        /// <param name="PersonId"></param>
        /// <returns></returns>
        public static bool GetInEntryView(Guid PersonId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<Guid> lstAbits = (from ab in context.Abiturient
                                       where ab.PersonId == PersonId
                                       select ab.Id).ToList();

                int cntProt = (from ph in context.extEntryView
                               where lstAbits.Contains(ph.AbiturientId)
                               select ph.AbiturientId).Count();

                if (cntProt > 0)
                    return true;
                else
                    return false;
            }
        }

        /// <summary>
        /// Возвращает, есть ли человек в протоколе о допуске
        /// </summary>
        /// <param name="PersonId"></param>
        /// <returns></returns>
        public static bool GetInEnableProtocol(Guid PersonId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                List<Guid> lstAbits = (from ab in context.Abiturient
                                       where ab.PersonId == PersonId
                                       select ab.Id).ToList();

                int cntProt = (from ph in context.extProtocol
                               where ph.ProtocolTypeId == 1 && !ph.IsOld && !ph.Excluded && lstAbits.Contains(ph.AbiturientId)
                               select ph.AbiturientId).Count();
                if (cntProt > 0)
                    return true;
                else
                    return false;
            }
        }

        public static extPerson_EducationInfo GetPersonCurrentInfo(Guid PersonId)
        {
            IPersonCurrentInfoDataProvider provider;
            switch (MainClass.studyLevelGroupId)
            {
                case 1: { provider = new PersonCurrentEducationDocument_1K(); break; }
                case 2: { provider = new PersonCurrentEducationDocument_Mag(); break; }
                case 3: { provider = new PersonCurrentEducationDocument_SPO(); break; }
                case 4: { provider = new PersonCurrentEducationDocument_Aspirant(); break; }
                default: { provider = new PersonCurrentEducationDocument_1K(); break; }
            }

            return provider.GetPersonCurrentInfo(PersonId);
        }

        public static extPerson GetExtPerson(Guid PersonId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                return context.extPerson.Where(x => x.Id == PersonId).FirstOrDefault();
            }
        }
    }

    /// <summary>
    /// Основной интерфейс поставки данных об текущем образовании
    /// </summary>
    public interface IPersonCurrentInfoDataProvider
    {
        extPerson_EducationInfo GetPersonCurrentInfo(Guid PersonId);
    }
    /// <summary>
    /// Generiс-интерфейс для точной поставки данных об текущем образовании. Требует точного указания типа приёма
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public interface IPersonCurrentInfoDataProvider<T> : IPersonCurrentInfoDataProvider where T : Common_Classes.PriemTypeClass
    {
    }

    //Производные классы-поставщики. Каждый может описывать свою собственную логику получения данных в зависимости от типа приёма
    public class PersonCurrentEducationDocument_1K : IPersonCurrentInfoDataProvider<Common_Classes.PriemType1K>
    {
        public extPerson_EducationInfo GetPersonCurrentInfo(Guid PersonId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                return (from EI in context.extPerson_EducationInfo
                        join Curr in context.hlpPersonMaxEducationInfoId on EI.Id equals Curr.Id
                        select EI).FirstOrDefault();
            }
        }
    }
    public class PersonCurrentEducationDocument_Mag : IPersonCurrentInfoDataProvider<Common_Classes.PriemTypeMag>
    {
        public extPerson_EducationInfo GetPersonCurrentInfo(Guid PersonId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                return (from EI in context.extPerson_EducationInfo
                        join Curr in context.hlpPersonMaxEducationInfoId on EI.Id equals Curr.Id
                        select EI).FirstOrDefault();
            }
        }
    }
    public class PersonCurrentEducationDocument_SPO : IPersonCurrentInfoDataProvider<Common_Classes.PriemTypeSPO>
    {
        public extPerson_EducationInfo GetPersonCurrentInfo(Guid PersonId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                return (from EI in context.extPerson_EducationInfo
                        join Curr in context.hlpPersonMaxEducationInfoId on EI.Id equals Curr.Id
                        select EI).FirstOrDefault();
            }
        }
    }
    public class PersonCurrentEducationDocument_Aspirant : IPersonCurrentInfoDataProvider<Common_Classes.PriemTypeAspirant>
    {
        public extPerson_EducationInfo GetPersonCurrentInfo(Guid PersonId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                return (from EI in context.extPerson_EducationInfo
                        join Curr in context.hlpPersonMaxEducationInfoId on EI.Id equals Curr.Id
                        select EI).FirstOrDefault();
            }
        }
    }
}
