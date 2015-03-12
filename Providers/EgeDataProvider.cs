using EducServLib;
using System;
using System.Collections.Generic;
using System.Data.Objects;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Priem
{
    public static class EgeDataProvider
    {
        public static void SaveEgeFromEgeList(EgeList egeLst, Guid PersonId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                foreach (EgeCertificateClass cert in egeLst.EGEs.Keys)
                {
                    if (!CheckEgeCertificate(cert))
                    {
                        WinFormsServ.Error(string.Format("Свидетельство с номером {0} уже есть в базе, поэтому сохранено не будет!", cert.Doc));
                        continue;
                    }

                    SaveEgeCertificate(cert, egeLst.EGEs[cert], PersonId);
                }
            }
        }

        // проверка на отсутствие одинаковых свидетельств
        public static bool CheckEgeCertificate(EgeCertificateClass cert)
        {
            using (PriemEntities context = new PriemEntities())
            {
                int res = (from ec in context.EgeCertificate
                           where ec.Number == cert.Doc
                           select ec).Count();
                return res > 0;
            }
        }

        public static void SaveEgeCertificate(EgeCertificateClass cert, List<EgeMarkCert> mrks, Guid personId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                ObjectParameter ecId = new ObjectParameter("id", typeof(Guid));
                context.EgeCertificate_Insert(cert.Doc, cert.Tipograf, "20" + cert.Doc.Substring(cert.Doc.Length - 2, 2), personId, null, false, ecId);

                if (ecId.Value == null)
                    return;

                Guid certId = (Guid)ecId.Value;
                foreach (EgeMarkCert mark in mrks)
                {
                    int val;
                    if (!int.TryParse(mark.Value, out val))
                        continue;

                    int subj;
                    if (!int.TryParse(mark.Subject, out subj))
                        continue;

                    context.EgeMark_Insert(val, subj, certId, false, false);
                }
            }
        }

        public static bool GetIsMatchEgeNumber(string number)
        {
            string num = number.Trim();
            if (Regex.IsMatch(num, @"^\d{2}-\d{9}-(09|10|11|12)$") || string.IsNullOrEmpty(num))
                return true;
            else
                return false;
        }

        public static string GetEgeCertificateNumbers(Guid PersonId, Guid EntryId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                string egecert =
                    (from egeCertificate in context.EgeCertificate
                     join egeMark in context.EgeMark on egeCertificate.Id equals egeMark.EgeCertificateId
                     join egeToExam in context.EgeToExam on egeMark.EgeExamNameId equals egeToExam.EgeExamNameId
                     where egeCertificate.PersonId == PersonId
                            && egeToExam.ExamId == (from examInEntry in context.ExamInEntry
                                                    where examInEntry.EntryId == EntryId && examInEntry.IsProfil
                                                    select examInEntry.ExamId).FirstOrDefault()
                     select new { egeCertificate.Year, egeCertificate.Number }).ToList().Distinct()
                     .Select(x => string.IsNullOrEmpty(x.Number) ? "Сертификат " + x.Year + " г." : x.Number).DefaultIfEmpty("")
                     .Aggregate((x, tail) => x + ", " + tail);

                return egecert;
            }
        }
    }
}
