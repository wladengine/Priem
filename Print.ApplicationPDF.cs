using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Priem
{
    partial class Print
    {
        //1курс-магистратура ОСНОВНОЙ (AbitTypeId = 1)
        public static byte[] GetApplicationPDF(string dirPath, bool isMag, Guid PersonId)
        {
            List<byte[]> lstFiles = new List<byte[]>();
            List<byte[]> lstAppendixes = new List<byte[]>();

            List<ApplicationData> abitList = ApplicationDataProvider.GetAppData(PersonId);
            extPerson person = PersonDataProvider.GetExtPerson(PersonId);
            extPerson_EducationInfo PersonEduc = PersonDataProvider.GetPersonCurrentInfo(PersonId);

            MemoryStream ms = new MemoryStream();
            string dotName;

            if (isMag)//mag
                dotName = "ApplicationMag_page3.pdf";
            else
                dotName = "Application_page3.pdf";

            PdfReader pdfRd = GetAcrobatFileFromTemplate(dirPath + "\\" + dotName);
            PdfStamper pdfStm = GetPdfStamper(ref pdfRd, ref ms, false);
            AcroFields acrFlds = pdfStm.AcroFields;

            FillInnerPriorData(abitList, dirPath, person, isMag, ref lstFiles, ref lstAppendixes);

            if (person.HostelEduc)
                acrFlds.SetField("HostelEducYes", "1");
            else
                acrFlds.SetField("HostelEducNo", "1");

            if (abitList.Where(x => x.IsGosLine).Count() > 0)
                acrFlds.SetField("IsGosLine", "1");

            acrFlds.SetField("HostelAbitYes", person.HostelAbit ? "1" : "0");
            acrFlds.SetField("HostelAbitNo", person.HostelAbit ? "0" : "1");

            acrFlds.SetField("BirthDateYear", person.BirthDate.Year.ToString("D2"));
            acrFlds.SetField("BirthDateMonth", person.BirthDate.Month.ToString("D2"));
            acrFlds.SetField("BirthDateDay", person.BirthDate.Day.ToString());

            acrFlds.SetField("BirthPlace", person.BirthPlace);
            acrFlds.SetField("Male", person.Sex ? "1" : "0");
            acrFlds.SetField("Female", person.Sex ? "0" : "1");
            acrFlds.SetField("Nationality", person.NationalityName);
            acrFlds.SetField("PassportSeries", person.PassportSeries);
            acrFlds.SetField("PassportNumber", person.PassportNumber);

            //dd.MM.yyyy :12.05.2000
            string[] splitStr = GetSplittedStrings(person.PassportAuthor + " " + person.PassportDate.Value.ToString("dd.MM.yyyy"), 60, 70, 2);
            for (int i = 1; i <= 2; i++)
                acrFlds.SetField("PassportAuthor" + i, splitStr[i - 1]);

            if (person.HasRussianNationality)
                acrFlds.SetField("HasRussianNationalityYes", "1");
            else
                acrFlds.SetField("HasRussianNationalityNo", "1");

            string Address = string.Format("{0} {1}{2},", (person.Code) ?? "", (person.NationalityId == 1 ? (person.RegionName + ", ") ?? "" : person.CountryName + ", "), (person.City + ", ") ?? "") +
                string.Format("{0} {1} {2} {3}", person.Street ?? "", person.House == string.Empty ? "" : "дом " + person.House,
                person.Korpus == string.Empty ? "" : "корп. " + person.Korpus,
                person.Flat == string.Empty ? "" : "кв. " + person.Flat);

            splitStr = GetSplittedStrings(Address, 50, 70, 3);
            for (int i = 1; i <= 3; i++)
                acrFlds.SetField("Address" + i, splitStr[i - 1]);

            acrFlds.SetField("EnglishMark", person.EnglishMark.ToString());
            if (person.StartEnglish)
                acrFlds.SetField("chbEnglishYes", "1");
            else
                acrFlds.SetField("chbEnglishNo", "1");

            acrFlds.SetField("Phone", person.Phone);
            acrFlds.SetField("Email", person.Email);
            acrFlds.SetField("Mobiles", person.Mobiles);

            acrFlds.SetField("ExitYear", PersonEduc.SchoolExitYear.ToString());
            splitStr = GetSplittedStrings(PersonEduc.SchoolName ?? "", 50, 70, 2);
            for (int i = 1; i <= 2; i++)
                acrFlds.SetField("School" + i, splitStr[i - 1]);

            //только у магистров
            acrFlds.SetField("HEProfession", PersonEduc.HEProfession ?? "");
            acrFlds.SetField("Qualification", PersonEduc.HEQualification ?? "");

            acrFlds.SetField("Original", "0");
            acrFlds.SetField("Copy", "0");
            acrFlds.SetField("CountryEduc", PersonEduc.CountryEducName ?? "");
            acrFlds.SetField("Language", person.LanguageName ?? "");

            string extraPerson = person.PersonInfo ?? "";
            splitStr = GetSplittedStrings(extraPerson, 70, 70, 3);
            for (int i = 1; i <= 3; i++)
            {
                acrFlds.SetField("Parents" + i.ToString(), splitStr[i - 1]);
                acrFlds.SetField("ExtraParents" + i.ToString(), splitStr[i - 1]);
            }

            string Attestat = PersonEduc.SchoolTypeId == 1 ? ("аттестат серия " + (PersonEduc.AttestatSeries ?? "") + " №" + (PersonEduc.AttestatNum ?? "")) :
                    ("диплом серия " + (PersonEduc.DiplomSeries ?? "") + " №" + (PersonEduc.DiplomNum ?? ""));
            acrFlds.SetField("Attestat", Attestat);
            acrFlds.SetField("Extra", person.ExtraInfo ?? "");

            if (PersonEduc.IsEqual && PersonEduc.CountryEducId != 193)
            {
                acrFlds.SetField("IsEqual", "1");
                acrFlds.SetField("EqualSertificateNumber", PersonEduc.EqualDocumentNumber);
            }
            else
                acrFlds.SetField("NoEqual", "1");

            if (person.Privileges > 0)
                acrFlds.SetField("HasPrivileges", "1");

            if ((PersonEduc.SchoolTypeId == 1) || (isMag && PersonEduc.SchoolTypeId == 4 && (PersonEduc.HEQualification).ToLower().IndexOf("магист") < 0))
                acrFlds.SetField("NoEduc", "1");
            else
            {
                acrFlds.SetField("HasEduc", "1");
                acrFlds.SetField("HighEducation", PersonEduc.SchoolName);
            }

            if (!isMag)
                AddEgeInfo(ref acrFlds, PersonId);

            if (!string.IsNullOrEmpty(PersonEduc.SchoolName))
                acrFlds.SetField("chbSchoolFinished", "1");

            if (!string.IsNullOrEmpty(person.Stag))
            {
                acrFlds.SetField("HasStag", "1");
                acrFlds.SetField("WorkPlace", person.WorkPlace);
                acrFlds.SetField("Stag", person.Stag);
            }
            else
                acrFlds.SetField("NoStag", "1");

            int comInd = 1;
            foreach (var comission in abitList.Select(x => x.ComissionAddress).Distinct().ToList())
                acrFlds.SetField("Comission" + comInd++, comission.ToString());

            pdfStm.FormFlattening = true;
            pdfStm.Close();
            pdfRd.Close();

            lstFiles.Add(ms.ToArray());

            return MergePdfFiles(lstFiles.Union(lstAppendixes).ToList());
        }
        public static void AddEgeInfo(ref AcroFields acrFlds, Guid PersonId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                //EGE
                var exams =
                    context.extEgeMark.Where(x => x.PersonId == PersonId).Select(x => new
                    {
                        ExamName = x.EgeExamName,
                        MarkValue = x.Value,
                        x.Number
                    }).ToList();
                int egeCnt = 1;
                foreach (var ex in exams)
                {
                    acrFlds.SetField("TableName" + egeCnt, ex.ExamName);
                    acrFlds.SetField("TableValue" + egeCnt, ex.MarkValue.ToString());
                    acrFlds.SetField("TableNumber" + egeCnt, ex.Number);

                    if (egeCnt == 4)
                        break;
                    egeCnt++;
                }

                //VSEROS
                var OlympVseros = context.Olympiads.Where(x => x.Abiturient.PersonId == PersonId && x.OlympTypeId == 2)
                    .Select(x => new { x.OlympSubject.Name, x.DocumentDate, x.DocumentSeries, x.DocumentNumber }).Distinct().ToList();
                egeCnt = 1;
                foreach (var ex in OlympVseros)
                {
                    acrFlds.SetField("OlympVserosName" + egeCnt, ex.Name);
                    acrFlds.SetField("OlympVserosYear" + egeCnt, ex.DocumentDate.HasValue ? ex.DocumentDate.Value.Year.ToString() : "");
                    acrFlds.SetField("OlympVserosDiplom" + egeCnt, (ex.DocumentSeries + " " ?? "") + (ex.DocumentNumber ?? ""));

                    if (egeCnt == 2)
                        break;
                    egeCnt++;
                }

                //OTHEROLYMPS
                var OlympNoVseros = context.Olympiads.Where(x => x.Abiturient.PersonId == PersonId && x.OlympTypeId != 2)
                    .Select(x => new { x.OlympName.Name, OlympSubject = x.OlympSubject.Name, x.DocumentDate, x.DocumentSeries, x.DocumentNumber }).ToList();
                egeCnt = 1;
                foreach (var ex in OlympNoVseros)
                {
                    acrFlds.SetField("OlympName" + egeCnt, ex.Name + " (" + ex.OlympSubject + ")");
                    acrFlds.SetField("OlympYear" + egeCnt, ex.DocumentDate.HasValue ? ex.DocumentDate.Value.Year.ToString() : "");
                    acrFlds.SetField("OlympDiplom" + egeCnt, (ex.DocumentSeries + " " ?? "") + (ex.DocumentNumber ?? ""));

                    if (egeCnt == 2)
                        break;
                    egeCnt++;
                }


            }
        }
        private static void FillInnerPriorData(List<ApplicationData> abitList, string dirPath, extPerson person, bool isMag, ref List<byte[]> lstFiles, ref List<byte[]> lstAppendixes)
        {
            List<ShortAppcationDetails> abitProfileList = ApplicationDataProvider.GetAppProfileList(person.Id);
            List<ShortAppcation> lstApps =
                abitList.Select(x => new ShortAppcation()
                {
                    ApplicationId = x.ApplicationId,
                    LicenseProgramName = x.LicenseProgramCode + " " + x.LicenseProgramName,
                    ObrazProgramName = x.ObrazProgramCrypt + " " + x.ObrazProgramName,
                    ProfileName = x.ProfileName,
                    Priority = x.Priority ?? 1,
                    StudyBasisId = x.StudyBasisId,
                    StudyFormId = x.StudyFormId,
                    HasInnerPriorities = abitProfileList.Where(y => y.ApplicationId == x.ApplicationId).Count() > 0,
                }).ToList();

            int incrmtr = 1;
            for (int i = 0; i < lstApps.Count; i++)
            {
                if (lstApps[i].HasInnerPriorities) //если есть профили
                {
                    lstApps[i].InnerPrioritiesNum = incrmtr; //то пишем об этом

                    //и сразу же создаём приложение с описанием - потом приложим
                    lstAppendixes.Add(GetApplicationPDF_Appendix_Mag(abitProfileList.Where(x => x.ApplicationId == lstApps[i].ApplicationId).ToList(), lstApps[i].LicenseProgramName, person.FIO, dirPath, incrmtr));
                    incrmtr++;
                }
            }

            //на 1 странице размещаем первые 3 конкурса
            List<ShortAppcation> lstAppsFirst = lstApps.Take(3).ToList();
            string code = (MainClass.iPriemYear % 100).ToString() + person.Num.ToString("D5");
            //добавляем первый файл
            lstFiles.Add(GetApplicationPDF_FirstPage(lstAppsFirst, lstApps, dirPath, isMag ? "ApplicationMag_page1.pdf" : "Application_page1.pdf", person.FIO, code, isMag));

            //остальные - по 4 на новую страницу
            int appcount = 3;
            while (appcount < lstApps.Count)
            {
                lstAppsFirst = new List<ShortAppcation>();
                for (int u = 0; u < 4; u++)
                {
                    if (lstApps.Count > appcount)
                        lstAppsFirst.Add(lstApps[appcount]);
                    else
                        break;
                    appcount++;
                }

                lstFiles.Add(GetApplicationPDF_NextPage(lstAppsFirst, lstApps, dirPath, "ApplicationMag_page2.pdf", person.FIO));
            }
        }

        public static byte[] GetApplicationPDF_Appendix_Mag(List<ShortAppcationDetails> lst, string LicenseProgramName, string FIO, string dirPath, int Num)
        {
            MemoryStream ms = new MemoryStream();
            string dotName = "PriorityProfiles_Mag2014.pdf";

            PdfReader pdfRd = GetAcrobatFileFromTemplate(dirPath + "\\" + dotName);
            PdfStamper pdfStm = GetPdfStamper(ref pdfRd, ref ms, false);

            AcroFields acrFlds = pdfStm.AcroFields;
            acrFlds.SetField("Num", Num.ToString());
            acrFlds.SetField("FIO", FIO);

            acrFlds.SetField("ObrazProgramHead", lst.First().ObrazProgramName);
            acrFlds.SetField("LicenseProgram", LicenseProgramName);
            acrFlds.SetField("ObrazProgram", lst.First().ObrazProgramName);
            int rwind = 1;
            foreach (var p in lst.Select(x => new { x.ProfileName, x.Priority }).Distinct().OrderBy(x => x.Priority))
                acrFlds.SetField("Profile" + rwind++, p.ProfileName);

            pdfStm.FormFlattening = true;
            pdfStm.Close();
            pdfRd.Close();

            return ms.ToArray();
        }

        public static byte[] GetApplicationPDF_FirstPage(List<ShortAppcation> lst, List<ShortAppcation> lstFullSource, string dirPath, string dotName, string FIO, string regNum, bool isMag)
        {
            MemoryStream ms = new MemoryStream();
            PdfReader pdfRd = GetAcrobatFileFromTemplate(dirPath + "\\" + dotName);
            PdfStamper pdfStm = GetPdfStamper(ref pdfRd, ref ms, false);

            AcroFields acrFlds = pdfStm.AcroFields;
            acrFlds.SetField("FIO", FIO);
            acrFlds.SetField("RegNum", regNum);

            int rwind = 1;
            foreach (var p in lst.OrderBy(x => x.Priority))
            {
                acrFlds.SetField("Priority" + rwind, p.Priority.ToString());
                acrFlds.SetField("Profession" + rwind, p.LicenseProgramName);
                acrFlds.SetField("ObrazProgram" + rwind, p.ObrazProgramName);
                acrFlds.SetField("Specialization" + rwind, p.HasInnerPriorities ? "Приложение к заявлению № " + p.InnerPrioritiesNum : p.ProfileName);
                acrFlds.SetField("StudyForm" + p.StudyFormId.ToString() + rwind.ToString(), "1");
                acrFlds.SetField("StudyBasis" + p.StudyBasisId.ToString() + rwind.ToString(), "1");

                if (lstFullSource.Where(x => x.LicenseProgramName == p.LicenseProgramName && x.ObrazProgramName == p.ObrazProgramName && x.ProfileName == p.ProfileName && x.StudyFormId == p.StudyFormId).Count() > 1)
                    acrFlds.SetField("IsPriority" + rwind, "1");

                rwind++;
            }

            pdfStm.FormFlattening = true;
            pdfStm.Close();
            pdfRd.Close();

            return ms.ToArray();
        }
        public static byte[] GetApplicationPDF_NextPage(List<ShortAppcation> lst, List<ShortAppcation> lstFullSource, string dirPath, string dotName, string FIO)
        {
            MemoryStream ms = new MemoryStream();
            PdfReader pdfRd = GetAcrobatFileFromTemplate(dirPath + "\\" + dotName);
            PdfStamper pdfStm = GetPdfStamper(ref pdfRd, ref ms, false);

            AcroFields acrFlds = pdfStm.AcroFields;
            int rwind = 1;
            foreach (var p in lst.OrderBy(x => x.Priority))
            {
                acrFlds.SetField("Priority" + rwind, p.Priority.ToString());
                acrFlds.SetField("Profession" + rwind, p.LicenseProgramName);
                acrFlds.SetField("ObrazProgram" + rwind, p.ObrazProgramName);
                acrFlds.SetField("Specialization" + rwind, p.HasInnerPriorities ? "Приложение к заявлению № " + p.InnerPrioritiesNum : p.ProfileName);
                acrFlds.SetField("StudyForm" + p.StudyFormId.ToString() + rwind.ToString(), "1");
                acrFlds.SetField("StudyBasis" + p.StudyBasisId.ToString() + rwind.ToString(), "1");

                if (lstFullSource.Where(x => x.LicenseProgramName == p.LicenseProgramName && x.ObrazProgramName == p.ObrazProgramName && x.ProfileName == p.ProfileName && x.StudyFormId == p.StudyFormId).Count() > 1)
                    acrFlds.SetField("IsPriority" + rwind, "1");

                rwind++;
            }

            pdfStm.FormFlattening = true;
            pdfStm.Close();
            pdfRd.Close();

            return ms.ToArray();
        }

        public static PdfReader GetAcrobatFileFromTemplate(string templateFile)
        {
            byte[] templateBytes;
            using (FileStream fs = new FileStream(templateFile, FileMode.Open, FileAccess.Read))
            {
                templateBytes = new byte[fs.Length];
                fs.Read(templateBytes, 0, templateBytes.Length);
            }

            return new PdfReader(templateBytes);
        }
        public static PdfStamper GetPdfStamper(ref PdfReader pdfRd, ref MemoryStream ms, bool setEncryption)
        {
            PdfStamper pdfStm = new PdfStamper(pdfRd, ms);
            if (setEncryption)
                pdfStm.SetEncryption(PdfWriter.STRENGTH128BITS, "", "", PdfWriter.ALLOW_SCREENREADERS | PdfWriter.ALLOW_PRINTING | PdfWriter.AllowPrinting);

            return pdfStm;
        }
    }
}
