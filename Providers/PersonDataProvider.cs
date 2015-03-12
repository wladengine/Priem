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
                    throw new Exception("Записей не найдено");

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
                context.Person_EducationInfo_insert(ED.PersonId, ED.IsExcellent, ED.SchoolCity, ED.SchoolTypeId, ED.SchoolName,
                    ED.SchoolNum, ED.SchoolExitYear, ED.SchoolAVG, ED.CountryEducId, ED.RegionEducId, ED.IsEqual,
                    ED.AttestatSeries, ED.AttestatNum, ED.DiplomSeries, ED.DiplomNum, ED.HighEducation,
                    ED.HEProfession, ED.HEQualification, ED.HEEntryYear, ED.HEExitYear, ED.HEStudyFormId, ED.HEWork, idParam);
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
    }
}
