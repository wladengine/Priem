using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using System.Data;
using System.Data.Objects;
using WordOut;
using iTextSharp.text;
using iTextSharp.text.pdf;

using EducServLib;

namespace Priem
{
    public class Print
    {
        public static void PrintHostelDirection(Guid? persId, bool forPrint, string savePath)
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    extPerson person = (from per in context.extPerson
                                        where per.Id == persId
                                        select per).FirstOrDefault();

                    FileStream fileS = null;
                    using (FileStream fs = new FileStream(string.Format(@"{0}\HostelDirection.pdf", MainClass.dirTemplates), FileMode.Open, FileAccess.Read))
                    {

                        byte[] bytes = new byte[fs.Length];
                        fs.Read(bytes, 0, bytes.Length);
                        fs.Close();

                        PdfReader pdfRd = new PdfReader(bytes);

                        try
                        {
                            fileS = new FileStream(string.Format(savePath), FileMode.Create);
                        }
                        catch
                        {
                            if (fileS != null)
                                fileS.Dispose();
                            WinFormsServ.Error("Пожалуйста, закройте открытые файлы pdf");
                            return;
                        }


                        PdfStamper pdfStm = new PdfStamper(pdfRd, fileS);
                        pdfStm.SetEncryption(PdfWriter.STRENGTH128BITS, "", "",
        PdfWriter.ALLOW_SCREENREADERS | PdfWriter.ALLOW_PRINTING |
        PdfWriter.AllowPrinting);
                        AcroFields acrFlds = pdfStm.AcroFields;

                        acrFlds.SetField("Surname", person.Surname);
                        acrFlds.SetField("Name", person.Name);
                        acrFlds.SetField("LastName", person.SecondName);

                        acrFlds.SetField("Faculty", person.HostelFacultyAcr);
                        acrFlds.SetField("Nationality", person.NationalityName);
                        acrFlds.SetField("Country", person.CountryName);

                        acrFlds.SetField("Male", person.Sex ? "0" : "1");
                        acrFlds.SetField("Female", person.Sex ? "1" : "0");

                        pdfStm.FormFlattening = true;
                        pdfStm.Close();
                        pdfRd.Close();

                        Process pr = new Process();
                        if (forPrint)
                        {
                            pr.StartInfo.Verb = "Print";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                        else
                        {
                            pr.StartInfo.Verb = "Open";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                    }
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
        }

        public static void PrintExamPass(Guid? persId, string savePath, bool forPrint)
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    extPerson person = (from per in context.extPerson
                                        where per.Id == persId
                                        select per).FirstOrDefault();

                    FileStream fileS = null;

                    using (FileStream fs = new FileStream(string.Format(@"{0}\ExamPass.pdf", MainClass.dirTemplates), FileMode.Open, FileAccess.Read))
                    {
                        byte[] bytes = new byte[fs.Length];
                        fs.Read(bytes, 0, bytes.Length);
                        fs.Close();

                        PdfReader pdfRd = new PdfReader(bytes);

                        try
                        {
                            fileS = new FileStream(string.Format(savePath), FileMode.Create);
                        }
                        catch
                        {
                            if (fileS != null)
                                fileS.Dispose();
                            WinFormsServ.Error("Пожалуйста, закройте открытые файлы pdf");
                            return;
                        }


                        PdfStamper pdfStm = new PdfStamper(pdfRd, fileS);
                        pdfStm.SetEncryption(PdfWriter.STRENGTH128BITS, "", "",
        PdfWriter.ALLOW_SCREENREADERS | PdfWriter.ALLOW_PRINTING |
        PdfWriter.AllowPrinting);
                        AcroFields acrFlds = pdfStm.AcroFields;

                        Barcode128 barcode = new Barcode128();
                        barcode.Code = person.PersonNum;

                        PdfContentByte cb = pdfStm.GetOverContent(1);

                        iTextSharp.text.Image img = barcode.CreateImageWithBarcode(cb, null, null);
                        img.SetAbsolutePosition(135, 565);
                        cb.AddImage(img);

                        acrFlds.SetField("Surname", person.Surname);
                        acrFlds.SetField("Name", person.Name);
                        acrFlds.SetField("LastName", person.SecondName);

                        acrFlds.SetField("Birth", person.BirthDate.ToShortDateString());
                        acrFlds.SetField("PassportSeries", person.PassportSeries + " " + person.PassportNumber);

                        acrFlds.SetField("chbMale", person.Sex ? "0" : "1");
                        acrFlds.SetField("chbFemale", person.Sex ? "1" : "0");


                        pdfStm.FormFlattening = true;
                        pdfStm.Close();
                        pdfRd.Close();

                        Process pr = new Process();
                        if (forPrint)
                        {
                            pr.StartInfo.Verb = "Print";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                        else
                        {
                            pr.StartInfo.Verb = "Open";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                    }
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
        }

        public static void PrintExamListWord(Guid? abitId, bool forPrint)
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    extAbit abit = (from ab in context.extAbit
                                    where ab.Id == abitId
                                    select ab).FirstOrDefault();

                    extPerson person = (from per in context.extPerson
                                        where per.Id == abit.PersonId
                                        select per).FirstOrDefault();

                    WordDoc wd = new WordDoc(string.Format(@"{0}\ExamSheet.dot", MainClass.dirTemplates), !forPrint);
                    TableDoc td = wd.Tables[0];

                    td[0, 0] = abit.FacultyName;
                    td[0, 1] = abit.LicenseProgramName;
                    td[0, 2] = abit.ProfileName;
                    td[1, 1] = MainClass.sPriemYear;
                    td[1, 0] = abit.StudyBasisName.Substring(0, 1).ToUpper() + abit.StudyFormOldName.Substring(0, 1).ToUpper();
                    td[0, 10] = person.Surname;
                    td[0, 11] = person.Name;
                    td[0, 12] = person.SecondName;

                    td[2, 13] = abit.RegNum;
                    td[1, 14] = abit.FacultyAcr;
                    td[1, 10] = person.PassportSeries + "   " + person.PassportNumber;

                    // экзамены!!! 
                    int row = 4;
                    IEnumerable<extExamInEntry> exams = from ex in context.extExamInEntry
                                                        where ex.EntryId == abit.EntryId
                                                        orderby ex.ExamName
                                                        select ex;

                    foreach (extExamInEntry ex in exams)
                    {
                        string sItem = ex.ExamName;
                        if (sItem.Contains("ностран") && MainClass.IsFilologFac())
                            sItem += string.Format(" ({0})", abit.LanguageName);

                        string mark = (from mrk in context.qMark
                                       where mrk.AbiturientId == abit.Id && mrk.ExamInEntryId == ex.Id
                                       select mrk.Value).FirstOrDefault().ToString();

                        td[0, row] = sItem;
                        td[1, row] = mark;
                        row++;
                    }

                    if (forPrint)
                    {
                        wd.Print();
                        wd.Close();
                    }
                }
            }
            catch (WordException we)
            {
                WinFormsServ.Error(we.Message);
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
        }

        public static void PrintExamList(Guid? abitId, bool forPrint, string savePath)
        {
            FileStream fileS = null;

            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    extAbit abit = (from ab in context.extAbit
                                    where ab.Id == abitId
                                    select ab).FirstOrDefault();

                    extPerson person = (from per in context.extPerson
                                        where per.Id == abit.PersonId
                                        select per).FirstOrDefault();

                    using (FileStream fs = new FileStream(string.Format(@"{0}\ExamList.pdf", MainClass.dirTemplates), FileMode.Open, FileAccess.Read))
                    {

                        byte[] bytes = new byte[fs.Length];
                        fs.Read(bytes, 0, bytes.Length);
                        fs.Close();

                        PdfReader pdfRd = new PdfReader(bytes);

                        try
                        {
                            fileS = new FileStream(string.Format(savePath), FileMode.Create);
                        }
                        catch
                        {
                            if (fileS != null)
                                fileS.Dispose();
                            WinFormsServ.Error("Пожалуйста, закройте открытые файлы pdf");
                            return;
                        }

                        PdfStamper pdfStm = new PdfStamper(pdfRd, fileS);
                        AcroFields acrFlds = pdfStm.AcroFields;

                        Barcode128 barcode = new Barcode128();
                        barcode.Code = abit.PersonNum + @"\" + abit.RegNum;

                        PdfContentByte cb = pdfStm.GetOverContent(1);

                        iTextSharp.text.Image img = barcode.CreateImageWithBarcode(cb, null, null);
                        img.SetAbsolutePosition(15, 65);
                        cb.AddImage(img);

                        acrFlds.SetField("Faculty", abit.FacultyName);
                        acrFlds.SetField("Profession", abit.LicenseProgramName);
                        acrFlds.SetField("Specialization", abit.ProfileName);
                        acrFlds.SetField("Year", MainClass.sPriemYear);
                        acrFlds.SetField("Study", abit.StudyBasisName.Substring(0, 1).ToUpper() + abit.StudyFormOldName.Substring(0, 1).ToUpper());

                        acrFlds.SetField("Surname", person.Surname);
                        acrFlds.SetField("Name", person.Name);
                        acrFlds.SetField("SecondName", person.SecondName);
                        acrFlds.SetField("RegNumber", abit.RegNum);

                        acrFlds.SetField("FacultyAcr", abit.FacultyAcr);
                        acrFlds.SetField("Passport", person.PassportSeries + "   " + person.PassportNumber);

                        // экзамены!!! 
                        int i = 1;
                        IEnumerable<extExamInEntry> exams = from ex in context.extExamInEntry
                                                            where ex.EntryId == abit.EntryId
                                                            orderby ex.ExamName
                                                            select ex;

                        foreach (extExamInEntry ex in exams)
                        {
                            string sItem = ex.ExamName;
                            if (sItem.Contains("ностран") && MainClass.IsFilologFac())
                                sItem += string.Format(" ({0})", abit.LanguageName);

                            string mark = (from mrk in context.qMark
                                           where mrk.AbiturientId == abit.Id && mrk.ExamInEntryId == ex.Id
                                           select mrk.Value).FirstOrDefault().ToString();

                            acrFlds.SetField("Exam" + i, sItem);
                            acrFlds.SetField("Mark" + i, mark);
                            i++;
                        }

                        pdfStm.FormFlattening = true;
                        pdfStm.Close();
                        pdfRd.Close();

                        fileS.Close();

                        Process pr = new Process();
                        if (forPrint)
                        {
                            pr.StartInfo.Verb = "Print";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                        else
                        {
                            pr.StartInfo.Verb = "Open";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                    }
                }
            }

            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
            finally
            {
                if (fileS != null)
                    fileS.Dispose();
            }
        }

        public static void PrintSprav(Guid? abitId, bool forPrint)
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    extAbit abit = (from ab in context.extAbit
                                    where ab.Id == abitId
                                    select ab).FirstOrDefault();

                    extPerson person = (from per in context.extPerson
                                        where per.Id == abit.PersonId
                                        select per).FirstOrDefault();

                    WordDoc wd = new WordDoc(string.Format(@"{0}\Spravka.dot", MainClass.dirTemplates), !forPrint);
                    TableDoc td = wd.Tables[0];

                    string sForm;

                    if (abit.StudyFormId == 1)
                        sForm = "дневную форму обучения";
                    else if (abit.StudyFormId == 2)
                        sForm = "вечернюю форму обучения";
                    else
                        sForm = "заочную форму обучения";

                    wd.Fields["Section"].Text = sForm;

                    string vinFac = (from f in context.qFaculty
                                     where f.Id == abit.FacultyId
                                     select (f.VinName == null ? "на " + f.Name : f.VinName)).FirstOrDefault().ToLower();

                    wd.SetFields("Faculty", vinFac);
                    wd.SetFields("FIO", person.FIO);
                    wd.SetFields("Profession", abit.LicenseProgramName);

                    // оценки!!

                    IEnumerable<qMark> marks = from mrk in context.qMark
                                               where mrk.AbiturientId == abit.Id
                                               select mrk;


                    string query = string.Format("SELECT qMark.Value, qMark.PassDate, extExamInProgram.ExamName as Name FROM (qMark INNER JOIN extExamInProgram ON qMark.ExamInProgramId = extExamInProgram.Id) INNER JOIN qAbiturient ON qMark.AbiturientId = qAbiturient.Id WHERE qAbiturient.Id = '{0}'", abitId);

                    int i = 1;
                    foreach (qMark m in marks)
                    {
                        td[0, i] = i.ToString();
                        td[1, i] = m.ExamName;
                        td[2, i] = m.PassDate.Value.ToShortDateString();
                        if (m.Value == 0 || m.Value == 1)
                            td[3, i] = MarkClass.MarkProp(m.Value.ToString());
                        else
                            td[3, i] = m.Value.ToString();
                        td.AddRow(1);
                        i++;
                    }
                    td.DeleteLastRow();

                    if (forPrint)
                    {
                        wd.Print();
                        wd.Close();
                    }
                }
            }
            catch (WordException we)
            {
                WinFormsServ.Error(we.Message);
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
        }

        public static void PrintStikerOne(Guid? abitId, bool forPrint)
        {
            string dotName;

            if (MainClass.dbType == PriemType.PriemMag)
                dotName = "StikerOneMag";
            else
                dotName = "StikerOne";

            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    //AbiturientClass abit = AbiturientClass.GetInstanceFromDBForPrint(abitId);
                    var abit = context.extAbit.Where(x => x.Id == abitId).First();
                    //PersonClass person = PersonClass.GetInstanceFromDBForPrint(abit.PersonId);
                    var person = context.extPerson.Where(x => x.Id == abit.PersonId).First();
                    var currEduc = context.extPerson_EducationInfo_Current.Where(x => x.PersonId == abit.PersonId).FirstOrDefault();

                    WordDoc wd = new WordDoc(string.Format(@"{0}\{1}.dot", MainClass.dirTemplates, dotName), !forPrint);

                    wd.SetFields("Faculty", abit.FacultyName);
                    wd.SetFields("Num", abit.PersonNum + @"\" + abit.RegNum);
                    wd.SetFields("Surname", person.Surname);
                    wd.SetFields("Name", person.Name);
                    wd.SetFields("SecondName", person.SecondName);
                    wd.SetFields("Profession", "(" + abit.LicenseProgramCode + ") " + abit.LicenseProgramName + ", " + abit.ObrazProgramName);
                    wd.SetFields("Specialization", abit.ProfileName);
                    wd.SetFields("Citizen", person.NationalityName);
                    wd.SetFields("Phone", person.Phone + "; " + person.Mobiles);
                    wd.SetFields("Email", person.Email);

                    for (int i = 1; i < 3; i++)
                    {
                        if (i != abit.StudyFormId)
                            wd.Shapes["StudyForm" + i].Delete();
                    }

                    for (int i = 1; i < 3; i++)
                    {
                        if (i != abit.StudyBasisId)
                            wd.Shapes["StudyBasis" + i].Delete();
                    }

                    wd.Shapes["Comp1"].Visible = false;
                    wd.Shapes["Comp2"].Visible = false;
                    wd.Shapes["Comp3"].Visible = false;
                    wd.Shapes["Comp4"].Visible = false;
                    wd.Shapes["Comp5"].Visible = false;
                    wd.Shapes["Comp6"].Visible = false;

                    wd.Shapes["Comp" + abit.CompetitionId.ToString()].Visible = true;
                    wd.Shapes["HasAssignToHostel"].Visible = person.HasAssignToHostel ?? false;

                    if (abit.CompetitionId == 6 && abit.OtherCompetitionId.HasValue)
                        wd.Shapes["Comp" + abit.CompetitionId.ToString()].Visible = true;

                    if (MainClass.dbType != PriemType.PriemMag)
                    {
                        string sPrevYear = DateTime.Now.AddYears(-1).Year.ToString();
                        string sCurrYear = DateTime.Now.Year.ToString();
                        string egePrevYear = context.EgeCertificate.Where(x => x.PersonId == person.Id && x.Year == sPrevYear).Select(x => x.Number).FirstOrDefault();
                        string egeCurYear = context.EgeCertificate.Where(x => x.PersonId == person.Id && x.Year == sCurrYear).Select(x => x.Number).FirstOrDefault();

                        wd.SetFields("EgeNamePrevYear", egePrevYear);
                        wd.SetFields("EgeNameCurYear", egeCurYear);

                        int j = 1;

                        using (PriemEntities ctx = new PriemEntities())
                        {
                            var lst = (from olympiads in ctx.Olympiads
                                       join olympValue in ctx.OlympValue on olympiads.OlympValueId equals olympValue.Id into olympValue2
                                       from olympValue in olympValue2.DefaultIfEmpty()
                                       join olympSubject in ctx.OlympSubject on olympiads.OlympSubjectId equals olympSubject.Id into olympSubject2
                                       from olympSubject in olympSubject2.DefaultIfEmpty()
                                       join olympType in ctx.OlympType on olympiads.OlympTypeId equals olympType.Id into eolympType2
                                       from olympType in eolympType2.DefaultIfEmpty()
                                       where olympiads.AbiturientId == abitId
                                       select new
                                       {
                                           Id = olympiads.Id,
                                           Тип = olympType.Name,
                                           Предмет = olympSubject.Name,
                                           OlympValueId = olympValue.Id,
                                           Степень = olympValue.Name
                                       }).ToList().Distinct();

                            foreach (var v in lst)
                            {
                                wd.SetFields("Level" + j, v.Тип);
                                wd.SetFields("Value" + j, v.Степень);
                                wd.SetFields("Subject" + j, v.Предмет);
                                j++;
                            }
                        }
                    }
                    else
                        if (currEduc.DiplomSeries != "" || currEduc.DiplomNum != "")
                            wd.SetFields("DocEduc", string.Format("диплом серия {0} № {1}", currEduc.DiplomSeries, currEduc.DiplomNum));

                    if (forPrint)
                    {
                        wd.Print();
                        wd.Close();
                    }
                }
            }
            catch (WordException we)
            {
                WinFormsServ.Error(we.Message);
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
        }

        public static void PrintStikerAll(Guid? personId, Guid? abitId, bool forPrint)
        {
            string dotName;

            if (MainClass.dbType == PriemType.PriemMag)
                dotName = "StikerAllMag";
            else
                dotName = "StikerAll";

            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    var person = context.extPerson.Where(x => x.Id == personId).First();
                    var currEduc = context.extPerson_EducationInfo_Current.Where(x => x.PersonId == personId).FirstOrDefault();

                    WordDoc wd = new WordDoc(string.Format(@"{0}\{1}.dot", MainClass.dirTemplates, dotName), !forPrint);

                    wd.SetFields("Num", context.extAbit.Where(x => x.PersonId == person.Id).Select(x => x.PersonNum).First());
                    wd.SetFields("Surname", person.Surname);
                    wd.SetFields("Name", person.Name);
                    wd.SetFields("SecondName", person.SecondName);
                    wd.SetFields("Citizen", person.NationalityName);
                    wd.SetFields("Phone", person.Phone + "; " + person.Mobiles);
                    wd.SetFields("Email", person.Email);

                    wd.Shapes["Comp1"].Visible = false;
                    wd.Shapes["Comp2"].Visible = false;
                    wd.Shapes["Comp3"].Visible = false;
                    wd.Shapes["Comp4"].Visible = false;
                    wd.Shapes["Comp5"].Visible = false;
                    wd.Shapes["Comp6"].Visible = false;

                    wd.Shapes["HasAssignToHostel"].Visible = person.HasAssignToHostel.Value;

                    if (MainClass.dbType != PriemType.PriemMag)
                    {
                        string sPrevYear = DateTime.Now.AddYears(-1).Year.ToString();
                        string sCurrYear = DateTime.Now.Year.ToString();
                        string egePrevYear = context.EgeCertificate.Where(x => x.PersonId == person.Id && x.Year == sPrevYear).Select(x => x.Number).FirstOrDefault();
                        string egeCurYear = context.EgeCertificate.Where(x => x.PersonId == person.Id && x.Year == sCurrYear).Select(x => x.Number).FirstOrDefault();

                        wd.SetFields("EgeNamePrevYear", egePrevYear);
                        wd.SetFields("EgeNameCurYear", egeCurYear);

                        int j = 1;

                        using (PriemEntities ctx = new PriemEntities())
                        {
                            var lst = (from olympiads in ctx.Olympiads
                                       join olympValue in ctx.OlympValue on olympiads.OlympValueId equals olympValue.Id into olympValue2
                                       from olympValue in olympValue2.DefaultIfEmpty()
                                       join olympSubject in ctx.OlympSubject on olympiads.OlympSubjectId equals olympSubject.Id into olympSubject2
                                       from olympSubject in olympSubject2.DefaultIfEmpty()
                                       join olympType in ctx.OlympType on olympiads.OlympTypeId equals olympType.Id into eolympType2
                                       from olympType in eolympType2.DefaultIfEmpty()
                                       where olympiads.AbiturientId == abitId
                                       select new
                                       {
                                           Id = olympiads.Id,
                                           Тип = olympType.Name,
                                           Предмет = olympSubject.Name,
                                           OlympValueId = olympValue.Id,
                                           Степень = olympValue.Name
                                       }).ToList().Distinct();

                            foreach (var v in lst)
                            {
                                wd.SetFields("Level" + j, v.Тип);
                                wd.SetFields("Value" + j, v.Степень);
                                wd.SetFields("Subject" + j, v.Предмет);
                                j++;
                            }
                        }
                    }
                    else
                        if (currEduc != null && (currEduc.DiplomSeries != "" || currEduc.DiplomNum != ""))
                            wd.SetFields("DocEduc", string.Format("диплом серия {0} № {1}", currEduc.DiplomSeries, currEduc.DiplomNum));

                    if (forPrint)
                    {
                        wd.Print();
                        wd.Close();
                    }
                }
            }
            catch (WordException we)
            {
                WinFormsServ.Error(we.Message);
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
        }

        public static void PrintApplication(bool forPrint, string savePath, Guid? PersonId)
        {
            if (!PersonId.HasValue)
                return;

            using (FileStream fs = new FileStream(savePath, FileMode.Create))
            using (BinaryWriter bw = new BinaryWriter(fs))
            {
                bool isMag = MainClass.dbType == PriemType.PriemMag;
                byte[] buffer = GetApplicationPDF(MainClass.dirTemplates, isMag, PersonId.Value);
                fs.Write(buffer, 0, buffer.Length);
                fs.Flush();
                fs.Close();
            }

            System.Diagnostics.Process.Start(savePath);
        }

        //1курс-магистратура ОСНОВНОЙ (AbitTypeId = 1)
        public static byte[] GetApplicationPDF(string dirPath, bool isMag, Guid PersonId)
        {
            using (PriemEntities context = new PriemEntities())
            {
                var abitList = (from x in context.Abiturient
                                join Entry in context.Entry on x.EntryId equals Entry.Id
                                where Entry.StudyLevel.StudyLevelGroup.Id == MainClass.studyLevelGroupId
                                && x.IsGosLine == false
                                && x.PersonId == PersonId
                                && x.BackDoc == false
                                select new
                                {
                                    x.Id,
                                    x.PersonId,
                                    x.Barcode,
                                    Faculty = Entry.SP_Faculty.Name,
                                    Profession = Entry.SP_LicenseProgram.Name,
                                    ProfessionCode = Entry.SP_LicenseProgram.Code,
                                    ObrazProgram = Entry.StudyLevel.Acronym + "." + Entry.SP_ObrazProgram.Number + "." + MainClass.sPriemYear + " " + Entry.SP_ObrazProgram.Name,
                                    Specialization = Entry.SP_Profile.Name,
                                    Entry.StudyFormId,
                                    Entry.StudyForm.Name,
                                    Entry.StudyBasisId,
                                    EntryType = (Entry.StudyLevelId == 17 ? 2 : 1),
                                    Entry.StudyLevelId,
                                    x.Priority,
                                    x.IsGosLine,
                                    Entry.CommissionId,
                                    ComissionAddress = Entry.CommissionId
                                }).OrderBy(x => x.Priority).ToList();

                var abitProfileList = (from x in context.Abiturient
                                       join Ad in context.ApplicationDetails on x.Id equals Ad.ApplicationId
                                       join Entry in context.Entry on x.EntryId equals Entry.Id
                                       where Entry.StudyLevel.StudyLevelGroup.Id == MainClass.studyLevelGroupId
                                       && x.IsGosLine == false
                                       && x.PersonId == PersonId
                                       && x.BackDoc == false
                                       select new ShortAppcationDetails()
                                       {
                                           ApplicationId = x.Id,
                                           Priority = Ad.InnerEntryInEntryPriority,
                                           ObrazProgramName = ((Ad.InnerEntryInEntry.SP_ObrazProgram.SP_LicenseProgram.StudyLevel.Acronym + "." + Ad.InnerEntryInEntry.SP_ObrazProgram.Number + " ") ?? "") + Ad.InnerEntryInEntry.SP_ObrazProgram.Name,
                                           ProfileName = Ad.InnerEntryInEntry.SP_Profile.Name
                                       }).Distinct().ToList();

                var person = (from x in context.Person
                              where x.Id == PersonId
                              select new
                              {
                                  x.Surname,
                                  x.Name,
                                  x.SecondName,
                                  x.Barcode,
                                  x.Person_AdditionalInfo.HostelAbit,
                                  x.BirthDate,
                                  BirthPlace = x.BirthPlace ?? "",
                                  Sex = x.Sex,
                                  Nationality = x.Nationality.Name,
                                  Country = x.Person_Contacts.Country.Name,
                                  PassportType = x.PassportType.Name,
                                  x.PassportSeries,
                                  x.PassportNumber,
                                  x.PassportAuthor,
                                  x.PassportDate,
                                  x.Person_Contacts.City,
                                  Region = x.Person_Contacts.Region.Name,
                                  AddInfo = x.Person_AdditionalInfo.ExtraInfo,
                                  Parents = x.Person_AdditionalInfo.PersonInfo,
                                  x.Person_Contacts.Code,
                                  x.Person_Contacts.Street,
                                  x.Person_Contacts.House,
                                  x.Person_Contacts.Korpus,
                                  x.Person_Contacts.Flat,
                                  x.Person_Contacts.Phone,
                                  x.Person_Contacts.Email,
                                  x.Person_Contacts.Mobiles,
                                  x.Person_AdditionalInfo.StartEnglish,
                                  x.Person_AdditionalInfo.EnglishMark,
                                  Language = x.Person_AdditionalInfo.Language.Name,

                                  x.Person_EducationInfo.First().CountryEducId,
                                  CountryEduc = x.Person_EducationInfo.First().CountryEducId != null ? x.Person_EducationInfo.First().Country.Name : "",
                                  Qualification = x.Person_EducationInfo.First().HEQualification,
                                  x.Person_EducationInfo.First().SchoolTypeId,
                                  x.Person_EducationInfo.First().SchoolName,
                                  x.Person_EducationInfo.First().SchoolExitYear,
                                  x.Person_EducationInfo.First().IsEqual,
                                  x.Person_EducationInfo.First().EqualDocumentNumber,
                                  x.Person_EducationInfo.First().AttestatSeries,
                                  x.Person_EducationInfo.First().AttestatNum,
                                  EducationDocumentSeries = x.Person_EducationInfo.First().DiplomSeries,
                                  EducationDocumentNumber = x.Person_EducationInfo.First().DiplomNum,
                                  ProgramName = x.Person_EducationInfo.First().HEProfession,

                                  HasPrivileges = (x.Person_AdditionalInfo.Privileges ?? 0) > 0,
                                  x.Person_AdditionalInfo.HasTRKI,
                                  x.Person_AdditionalInfo.TRKICertificateNumber,
                                  x.Person_AdditionalInfo.HostelEduc,
                                  IsRussia = (x.Person_Contacts.CountryId == 1),
                                  x.HasRussianNationality,
                                  x.Person_AdditionalInfo.Stag,
                                  x.Person_AdditionalInfo.WorkPlace,
                                  x.Num
                              }).FirstOrDefault();

                MemoryStream ms = new MemoryStream();
                string dotName;

                if (isMag)//mag
                    dotName = "ApplicationMag_page3.pdf";
                else
                    dotName = "Application_page3.pdf";

                byte[] templateBytes;

                List<byte[]> lstFiles = new List<byte[]>();
                List<byte[]> lstAppendixes = new List<byte[]>();
                using (FileStream fs = new FileStream(dirPath + "\\" + dotName, FileMode.Open, FileAccess.Read))
                {
                    templateBytes = new byte[fs.Length];
                    fs.Read(templateBytes, 0, templateBytes.Length);
                }

                PdfReader pdfRd = new PdfReader(templateBytes);
                PdfStamper pdfStm = new PdfStamper(pdfRd, ms);
                //pdfStm.SetEncryption(PdfWriter.STRENGTH128BITS, "", "", PdfWriter.ALLOW_SCREENREADERS | PdfWriter.ALLOW_PRINTING | PdfWriter.AllowPrinting);
                AcroFields acrFlds = pdfStm.AcroFields;

                string FIO = ((person.Surname ?? "") + " " + (person.Name ?? "") + " " + (person.SecondName ?? "")).Trim();

                List<ShortAppcation> lstApps = abitList
                    .Select(x => new ShortAppcation()
                    {
                        ApplicationId = x.Id,
                        LicenseProgramName = x.ProfessionCode + " " + x.Profession,
                        ObrazProgramName = x.ObrazProgram,
                        ProfileName = x.Specialization,
                        Priority = x.Priority ?? 1,
                        StudyBasisId = x.StudyBasisId,
                        StudyFormId = x.StudyFormId,
                        HasInnerPriorities = abitProfileList.Where(y => y.ApplicationId == x.Id).Count() > 0,
                    }).ToList();
                int incrmtr = 1;
                for (int u = 0; u < lstApps.Count; u++)
                {
                    if (lstApps[u].HasInnerPriorities) //если есть профили
                    {
                        lstApps[u].InnerPrioritiesNum = incrmtr; //то пишем об этом
                        //и сразу же создаём приложение с описанием - потом приложим

                        if (isMag) //для магов всё просто
                        {
                            lstAppendixes.Add(GetApplicationPDF_ProfileAppendix_Mag(abitProfileList.Where(x => x.ApplicationId == lstApps[u].ApplicationId).ToList(), lstApps[u].LicenseProgramName, FIO, dirPath, incrmtr));
                            incrmtr++;
                        }
                        else //для перваков всё запутаннее
                        {   //сначала надо проверить, нет ли внутреннего разбиения по программам
                            //если есть, то для каждой программы сделать своё приложение, а затем уже для тех программ, где есть внутри профили доложить приложений с профилями
                            var profs = abitProfileList.Where(x => x.ApplicationId == lstApps[u].ApplicationId).Select(x => new ShortAppcationDetails() 
                            { 
                                ApplicationId = x.ApplicationId,  
                                ObrazProgramName = x.ObrazProgramName,
                                Priority = x.Priority,
                                ProfileName = x.ProfileName,
                            }).Distinct().ToList();
                            var OP = profs.Select(x => x.ObrazProgramName).Distinct().ToList();
                            if (OP.Count > 1)
                            {
                                lstAppendixes.Add(GetApplicationPDF_OPAppendix_1kurs(profs, lstApps[u].LicenseProgramName, FIO, dirPath, incrmtr));
                                incrmtr++;
                            }
                            foreach (var OP_name in OP)
                            {
                                var lstProfs = abitProfileList.Where(x => x.ApplicationId == lstApps[u].ApplicationId && x.ObrazProgramName == OP_name).Distinct().ToList();
                                if (lstProfs.Select(x => x.ProfileName).Distinct().Count() > 1)
                                {
                                    lstAppendixes.Add(GetApplicationPDF_ProfileAppendix_1kurs(lstProfs, lstApps[u].LicenseProgramName, FIO, dirPath, incrmtr));
                                    incrmtr++;
                                }
                            }
                        }
                    }
                }

                List<ShortAppcation> lstAppsFirst = new List<ShortAppcation>();
                for (int u = 0; u < 3; u++)
                {
                    if (lstApps.Count > u)
                        lstAppsFirst.Add(lstApps[u]);
                }

                string code =  (MainClass.iPriemYear % 100).ToString() + person.Num.ToString("D5");
                //добавляем первый файл
                lstFiles.Add(GetApplicationPDF_FirstPage(lstAppsFirst, lstApps, dirPath, isMag ? "ApplicationMag_page1.pdf" : "Application_page1.pdf", FIO, code, isMag));

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

                    lstFiles.Add(GetApplicationPDF_NextPage(lstAppsFirst, lstApps, dirPath, "ApplicationMag_page2.pdf", FIO));
                }


                if (person.HostelEduc)
                    acrFlds.SetField("HostelEducYes", "1");
                else
                    acrFlds.SetField("HostelEducNo", "1");

                if (abitList.Where(x => x.IsGosLine).Count() > 0)
                    acrFlds.SetField("IsGosLine", "1");

                acrFlds.SetField("HostelAbitYes", person.HostelAbit ? "1" : "0");
                acrFlds.SetField("HostelAbitNo", person.HostelAbit  ? "0" : "1");

                acrFlds.SetField("BirthDateYear", person.BirthDate.Year.ToString("D2"));
                acrFlds.SetField("BirthDateMonth", person.BirthDate.Month.ToString("D2"));
                acrFlds.SetField("BirthDateDay", person.BirthDate.Day.ToString());

                acrFlds.SetField("BirthPlace", person.BirthPlace);
                acrFlds.SetField("Male", person.Sex ? "1" : "0");
                acrFlds.SetField("Female", person.Sex ? "0" : "1");
                acrFlds.SetField("Nationality", person.Nationality);
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

                string Address = string.Format("{0} {1}{2},", (person.Code) ?? "", (person.IsRussia ? (person.Region + ", ") ?? "" : person.Country + ", "), (person.City + ", ") ?? "") +
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

                acrFlds.SetField("ExitYear", person.SchoolExitYear.ToString());
                splitStr = GetSplittedStrings(person.SchoolName ?? "", 50, 70, 2);
                for (int i = 1; i <= 2; i++)
                    acrFlds.SetField("School" + i, splitStr[i - 1]);

                //только у магистров
                acrFlds.SetField("HEProfession", person.ProgramName ?? "");
                acrFlds.SetField("Qualification", person.Qualification ?? "");

                acrFlds.SetField("Original", "0");
                acrFlds.SetField("Copy", "0");
                acrFlds.SetField("CountryEduc", person.CountryEduc ?? "");
                acrFlds.SetField("Language", person.Language ?? "");

                string extraPerson = person.Parents ?? "";
                splitStr = GetSplittedStrings(extraPerson, 70, 70, 3);
                for (int i = 1; i <= 3; i++)
                {
                    acrFlds.SetField("Parents" + i.ToString(), splitStr[i - 1]);
                    acrFlds.SetField("ExtraParents" + i.ToString(), splitStr[i - 1]);
                }

                string Attestat = person.SchoolTypeId == 1 ? ("аттестат серия " + (person.AttestatSeries ?? "") + " №" + (person.AttestatNum ?? "")) :
                        ("диплом серия " + (person.EducationDocumentSeries ?? "") + " №" + (person.EducationDocumentNumber ?? ""));
                acrFlds.SetField("Attestat", Attestat);
                acrFlds.SetField("Extra", person.AddInfo ?? "");

                if (person.IsEqual && person.CountryEducId != 193)
                {
                    acrFlds.SetField("IsEqual", "1");
                    acrFlds.SetField("EqualSertificateNumber", person.EqualDocumentNumber);
                }
                else
                {
                    acrFlds.SetField("NoEqual", "1");
                }

                if (person.HasPrivileges)
                    acrFlds.SetField("HasPrivileges", "1");

                if ((person.SchoolTypeId == 1) || (isMag && person.SchoolTypeId == 4 && (person.Qualification).ToLower().IndexOf("магист") < 0))
                    acrFlds.SetField("NoEduc", "1");
                else
                {
                    acrFlds.SetField("HasEduc", "1");
                    acrFlds.SetField("HighEducation", person.SchoolName);
                }

                if (!isMag)
                {
                    //EGE
                    var exams = context.extEgeMark.Where(x => x.PersonId == PersonId).Select(x => new
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

                    if (!string.IsNullOrEmpty(person.SchoolName))
                        acrFlds.SetField("chbSchoolFinished", "1");
                }
                
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
                {
                    acrFlds.SetField("Comission" + comInd++, comission.ToString());
                }

                context.SaveChanges();

                pdfStm.FormFlattening = true;
                pdfStm.Close();
                pdfRd.Close();

                lstFiles.Add(ms.ToArray());

                return MergePdfFiles(lstFiles.Union(lstAppendixes).ToList());
            }
        }

        public static byte[] GetApplicationPDF_ProfileAppendix_Mag(List<ShortAppcationDetails> lst, string LicenseProgramName, string FIO, string dirPath, int Num)
        {
            MemoryStream ms = new MemoryStream();
            string dotName = "PriorityProfiles_Mag2014.pdf";

            byte[] templateBytes;
            using (FileStream fs = new FileStream(dirPath + "\\" + dotName, FileMode.Open, FileAccess.Read))
            {
                templateBytes = new byte[fs.Length];
                fs.Read(templateBytes, 0, templateBytes.Length);
            }

            PdfReader pdfRd = new PdfReader(templateBytes);
            PdfStamper pdfStm = new PdfStamper(pdfRd, ms);
            //pdfStm.SetEncryption(PdfWriter.STRENGTH128BITS, "", "", PdfWriter.ALLOW_SCREENREADERS | PdfWriter.ALLOW_PRINTING | PdfWriter.AllowPrinting);
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
        public static byte[] GetApplicationPDF_OPAppendix_1kurs(List<ShortAppcationDetails> lst, string LicenseProgramName, string FIO, string dirPath, int Num)
        {
            MemoryStream ms = new MemoryStream();
            string dotName = "PriorityOP2014.pdf";

            byte[] templateBytes;
            using (FileStream fs = new FileStream(dirPath + "\\" + dotName, FileMode.Open, FileAccess.Read))
            {
                templateBytes = new byte[fs.Length];
                fs.Read(templateBytes, 0, templateBytes.Length);
            }

            PdfReader pdfRd = new PdfReader(templateBytes);
            PdfStamper pdfStm = new PdfStamper(pdfRd, ms);
            //pdfStm.SetEncryption(PdfWriter.STRENGTH128BITS, "", "", PdfWriter.ALLOW_SCREENREADERS | PdfWriter.ALLOW_PRINTING | PdfWriter.AllowPrinting);
            AcroFields acrFlds = pdfStm.AcroFields;
            acrFlds.SetField("Num", Num.ToString());
            acrFlds.SetField("FIO", FIO);

            acrFlds.SetField("LicenseProgram", LicenseProgramName);
            int rwind = 1;
            foreach (var p in lst.Select(x => new { x.ObrazProgramName, x.Priority }).Distinct().OrderBy(x => x.Priority))
                acrFlds.SetField("ObrazProgram" + rwind++, p.ObrazProgramName);
            
            pdfStm.FormFlattening = true;
            pdfStm.Close();
            pdfRd.Close();

            return ms.ToArray();
        }
        public static byte[] GetApplicationPDF_ProfileAppendix_1kurs(List<ShortAppcationDetails> lst, string LicenseProgramName, string FIO, string dirPath, int Num)
        {
            MemoryStream ms = new MemoryStream();
            string dotName = "PriorityProfiles2014.pdf";

            byte[] templateBytes;
            using (FileStream fs = new FileStream(dirPath + "\\" + dotName, FileMode.Open, FileAccess.Read))
            {
                templateBytes = new byte[fs.Length];
                fs.Read(templateBytes, 0, templateBytes.Length);
            }

            PdfReader pdfRd = new PdfReader(templateBytes);
            PdfStamper pdfStm = new PdfStamper(pdfRd, ms);
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

            byte[] templateBytes;
            using (FileStream fs = new FileStream(dirPath + "\\" + dotName, FileMode.Open, FileAccess.Read))
            {
                templateBytes = new byte[fs.Length];
                fs.Read(templateBytes, 0, templateBytes.Length);
            }

            PdfReader pdfRd = new PdfReader(templateBytes);
            PdfStamper pdfStm = new PdfStamper(pdfRd, ms);
            //pdfStm.SetEncryption(PdfWriter.STRENGTH128BITS, "", "", PdfWriter.ALLOW_SCREENREADERS | PdfWriter.ALLOW_PRINTING | PdfWriter.AllowPrinting);

            AcroFields acrFlds = pdfStm.AcroFields;
            acrFlds.SetField("FIO", FIO);
            
            //добавляем штрихкод
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

            byte[] templateBytes;
            using (FileStream fs = new FileStream(dirPath + "\\" + dotName, FileMode.Open, FileAccess.Read))
            {
                templateBytes = new byte[fs.Length];
                fs.Read(templateBytes, 0, templateBytes.Length);
            }

            PdfReader pdfRd = new PdfReader(templateBytes);
            PdfStamper pdfStm = new PdfStamper(pdfRd, ms);
            //pdfStm.SetEncryption(PdfWriter.STRENGTH128BITS, "", "", PdfWriter.ALLOW_SCREENREADERS | PdfWriter.ALLOW_PRINTING | PdfWriter.AllowPrinting);
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

        public static void PrintEnableProtocol(string protocolId, bool forPrint, string savePath)
        {
            FileStream fileS = null;
            try
            {
                Guid gProtocolId = Guid.Parse(protocolId);

                using (PriemEntities context = new PriemEntities())
                {
                    var info = ProtocolDataProvider.GetProtocolInfo(gProtocolId, 1);

                    string basis = string.Empty;
                    switch (info.StudyBasisId)
                    {
                        case 1:
                            basis = "Бюджетные места";
                            break;
                        case 2:
                            basis = "Места по договорам с оплатой стоимости обучения";
                            break;
                    }

                    Document document = new Document(PageSize.A4.Rotate(), 50, 50, 50, 50);

                    using (fileS = new FileStream(savePath, FileMode.Create))
                    {

                        BaseFont bfTimes = BaseFont.CreateFont(string.Format(@"{0}\times.ttf", MainClass.dirTemplates), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                        Font font = new Font(bfTimes, 10);

                        PdfWriter.GetInstance(document, fileS);
                        document.Open();

                        //HEADER
                        string header = string.Format(@"Форма обучения: {0}
    Условия обучения: {1}", info.StudyFormName, basis);

                        Paragraph p = new Paragraph(header, font);
                        document.Add(p);

                        float midStr = 13f;
                        p = new Paragraph(20f);
                        p.Add(new Phrase("ПРОТОКОЛ № ", new Font(bfTimes, 14, Font.BOLD)));
                        p.Add(new Phrase(info.Number, new Font(bfTimes, 18, Font.BOLD)));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        p = new Paragraph(midStr);
                        p.Add(new Phrase(@"заседания Приемной комиссии Санкт-Петербургского Государственного Университета
    о допуске к участию в конкурсе на основные образовательные программы ", new Font(bfTimes, 10, Font.BOLD)));

                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        //date
                        p = new Paragraph(midStr);
                        p.Add(new Paragraph(string.Format("от {0}", Util.GetDateString(info.Date, true, true)), new Font(bfTimes, 10, Font.BOLD)));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        string spec = "", currSpec = "";
                        PdfPTable curT = null;
                        int cnt = 0;

                        var lst = ProtocolDataProvider.GetProtocolData(gProtocolId);
                        foreach (var v in lst)
                        {
                            cnt++;

                            currSpec = v.Direction;
                            if (spec != currSpec)
                            {
                                spec = currSpec;
                                cnt = 1;

                                if (curT != null)
                                    document.Add(curT);

                                //Table
                                Table table = new Table(7);
                                table.Padding = 3;
                                table.Spacing = 0;
                                float[] headerwidths = { 5, 10, 30, 15, 20, 10, 10 };
                                table.Widths = headerwidths;
                                table.Width = 100;

                                PdfPTable t = new PdfPTable(7);
                                t.SetWidthPercentage(headerwidths, document.PageSize);
                                t.WidthPercentage = 100f;
                                t.SpacingBefore = 10f;
                                t.SpacingAfter = 10f;

                                t.HeaderRows = 2;

                                Phrase pra = new Phrase(string.Format("По направлению {0} ", currSpec), new Font(bfTimes, 10));

                                PdfPCell pcell = new PdfPCell(pra);
                                pcell.BorderWidth = 0;
                                pcell.Colspan = 7;
                                t.AddCell(pcell);

                                string[] headers = new string[]
                                    {
                                        "№ п/п",
                                        "Рег.номер",
                                        "ФАМИЛИЯ, ИМЯ, ОТЧЕСТВО",
                                        "Номер аттестата или диплома",
                                        "Номер сертификата ЕГЭ по профильному предмету",
                                        "Вид конкурса",
                                        "Примечания"
                                    };

                                foreach (string h in headers)
                                {
                                    PdfPCell cell = new PdfPCell();
                                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                    cell.AddElement(new Phrase(h, new Font(bfTimes, 10, Font.BOLD)));

                                    t.AddCell(cell);
                                }

                                curT = t;
                            }

                            string egecert = (from egeCertificate in context.EgeCertificate
                                              join egeMark in context.EgeMark on egeCertificate.Id equals egeMark.EgeCertificateId
                                              join egeToExam in context.EgeToExam on egeMark.EgeExamNameId equals egeToExam.EgeExamNameId
                                              where egeCertificate.PersonId == v.PersonId
                                              && egeToExam.ExamId == (from examInEntry in context.ExamInEntry
                                                                      where examInEntry.EntryId == v.EntryId && examInEntry.IsProfil
                                                                      select examInEntry.ExamId).FirstOrDefault()
                                              select egeCertificate.Number).FirstOrDefault();

                            curT.AddCell(new Phrase(cnt.ToString(), new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(v.RegNum, new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(v.FIO, new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(v.EducationDocument, new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(egecert, new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(v.CompetitionName, new Font(bfTimes, 10)));
                            curT.AddCell(new Phrase(v.Comment, new Font(bfTimes, 10)));
                        }

                        if (curT != null)
                            document.Add(curT);

                        //FOOTER
                        p = new Paragraph(30f);
                        p.KeepTogether = true;
                        p.Add(new Phrase("Ответственный секретарь Приемной комиссии СПбГУ____________________________________________________________", new Font(bfTimes, 10)));
                        document.Add(p);

                        p = new Paragraph();
                        p.Add(new Phrase("Заместитель начальника Управления по организации приема – советник проректора по направлениям___________________", new Font(bfTimes, 10)));
                        document.Add(p);

                        p = new Paragraph();
                        p.Add(new Phrase("Ответственный секретарь комиссии по приему документов_______________________________________________________", new Font(bfTimes, 10)));
                        document.Add(p);

                        document.Close();

                        Process pr = new Process();
                        if (forPrint)
                        {
                            pr.StartInfo.Verb = "Print";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                        else
                        {
                            pr.StartInfo.Verb = "Open";
                            pr.StartInfo.FileName = string.Format(savePath);
                            pr.Start();
                        }
                    }
                }
            }

            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
            finally
            {
                if (fileS != null)
                    fileS.Dispose();
            }
        }
        public static void PrintDisEnableProtocol(string protocolId, bool forPrint, string savePath)
        {
            FileStream fileS = null;
            try
            {
                Guid gProtocolId = Guid.Parse(protocolId);
                var protocolInfo = ProtocolDataProvider.GetProtocolInfo(gProtocolId, 2); //DisEnableProtocol

                string basis = string.Empty;
                switch (protocolInfo.StudyBasisId.ToString())
                {
                    case "1":
                        basis = "Бюджетные места";
                        break;
                    case "2":
                        basis = "Места по договорам с оплатой стоимости обучения";
                        break;
                }

                Document document = new Document(PageSize.A4.Rotate(), 50, 50, 50, 50);
                using (fileS = new FileStream(savePath, FileMode.Create))
                {
                    BaseFont bfTimes = BaseFont.CreateFont(string.Format(@"{0}\times.ttf", MainClass.dirTemplates), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    Font font = new Font(bfTimes, 10);

                    PdfWriter.GetInstance(document, fileS);
                    document.Open();

                    //HEADER
                    string header = string.Format(@"Форма обучения: {0}
Условия обучения: {1}", protocolInfo.StudyFormName, basis);

                    Paragraph p = new Paragraph(header, font);
                    document.Add(p);

                    float midStr = 13f;
                    p = new Paragraph(20f);
                    p.Add(new Phrase("ПРОТОКОЛ № ", new Font(bfTimes, 14, Font.BOLD)));
                    p.Add(new Phrase(protocolInfo.Number, new Font(bfTimes, 18, Font.BOLD)));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    p = new Paragraph(midStr);
                    p.Add(new Phrase(@"заседания Приемной комиссии Санкт-Петербургского Государственного Университета
об исключении из участия в конкурсе на основные образовательные программы ", new Font(bfTimes, 10, Font.BOLD)));

                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    //date
                    p = new Paragraph(midStr);
                    p.Add(new Paragraph(string.Format("от {0}", Util.GetDateString(protocolInfo.Date, true, true)), new Font(bfTimes, 10, Font.BOLD)));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    string spec = "", currSpec = "";
                    PdfPTable curT = null;
                    int cnt = 0;

                    var lst = ProtocolDataProvider.GetProtocolData(gProtocolId);
                    foreach (var v in lst)
                    {
                        cnt++;

                        currSpec = v.Direction;
                        if (spec != currSpec)
                        {
                            spec = currSpec;
                            cnt = 1;

                            if (curT != null)
                                document.Add(curT);

                            //Table
                            Table table = new Table(7);
                            table.Padding = 3;
                            table.Spacing = 0;
                            float[] headerwidths = { 5, 10, 30, 15, 20, 10, 10 };
                            table.Widths = headerwidths;
                            table.Width = 100;

                            PdfPTable t = new PdfPTable(7);
                            t.SetWidthPercentage(headerwidths, document.PageSize);
                            t.WidthPercentage = 100f;
                            t.SpacingBefore = 10f;
                            t.SpacingAfter = 10f;

                            t.HeaderRows = 2;

                            Phrase pra = new Phrase(string.Format("По направлению {0} ", currSpec), new Font(bfTimes, 10));

                            PdfPCell pcell = new PdfPCell(pra);
                            pcell.BorderWidth = 0;
                            pcell.Colspan = 7;
                            t.AddCell(pcell);

                            string[] headers = new string[]
                            {
                                "№ п/п",
                                "Рег.номер",
                                "ФАМИЛИЯ, ИМЯ, ОТЧЕСТВО",
                                "Номер аттестата или диплома",
                                "Номер сертификата ЕГЭ по профильному предмету",
                                "Вид конкурса",
                                "Примечания"
                            };

                            foreach (string h in headers)
                            {
                                PdfPCell cell = new PdfPCell();
                                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                cell.AddElement(new Phrase(h, new Font(bfTimes, 10, Font.BOLD)));

                                t.AddCell(cell);
                            }

                            curT = t;
                        }

                        string egecert = EgeDataProvider.GetEgeCertificateNumbers(v.PersonId, v.EntryId);

                        curT.AddCell(new Phrase(cnt.ToString(), new Font(bfTimes, 10)));
                        curT.AddCell(new Phrase(v.RegNum, new Font(bfTimes, 10)));
                        curT.AddCell(new Phrase(v.FIO, new Font(bfTimes, 10)));
                        curT.AddCell(new Phrase(v.EducationDocument, new Font(bfTimes, 10)));
                        curT.AddCell(new Phrase(egecert, new Font(bfTimes, 10)));
                        curT.AddCell(new Phrase(v.CompetitionName, new Font(bfTimes, 10)));
                        curT.AddCell(new Phrase(v.Comment, new Font(bfTimes, 10)));
                    }


                    if (curT != null)
                        document.Add(curT);

                    //FOOTER
                    p = new Paragraph(30f);
                    p.KeepTogether = true;
                    p.Add(new Phrase("Ответственный секретарь Приемной комиссии СПбГУ_______________________________________________________", new Font(bfTimes, 10)));
                    document.Add(p);

                    p = new Paragraph();
                    p.Add(new Phrase(@"Заместитель Ответственного секретаря Приемной 
комиссии  СПбГУ по группе основных образовательных программ_____________________________________________", new Font(bfTimes, 10)));
                    document.Add(p);

                    p = new Paragraph();
                    p.Add(new Phrase("Ответственный по приему на основную образовательную программу___________________________________________", new Font(bfTimes, 10)));
                    document.Add(p);

                    document.Close();

                    Process pr = new Process();
                    if (forPrint)
                    {
                        pr.StartInfo.Verb = "Print";
                        pr.StartInfo.FileName = string.Format(savePath);
                        pr.Start();
                    }
                    else
                    {
                        pr.StartInfo.Verb = "Open";
                        pr.StartInfo.FileName = string.Format(savePath);
                        pr.Start();
                    }
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
            finally
            {
                if (fileS != null)
                    fileS.Dispose();
            }
        }
        public static void PrintChangeCompCelProtocol(string protocolId, bool forPrint, string savePath)
        {
            FileStream fileS = null;
            try
            {
                Guid gProtocolId = Guid.Parse(protocolId);
                var info = ProtocolDataProvider.GetProtocolInfo(gProtocolId, 3);

                string basis = string.Empty;
                switch (info.StudyBasisId.ToString())
                {
                    case "1":
                        basis = "Бюджетные места";
                        break;
                    case "2":
                        basis = "Места по договорам с оплатой стоимости обучения";
                        break;
                }

                Document document = new Document(PageSize.A4.Rotate(), 50, 50, 50, 50);
                using (fileS = new FileStream(savePath, FileMode.Create))
                {

                    BaseFont bfTimes = BaseFont.CreateFont(string.Format(@"{0}\times.ttf", MainClass.dirTemplates), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    Font font = new Font(bfTimes, 10);

                    PdfWriter.GetInstance(document, fileS);
                    document.Open();

                    //HEADER
                    string header = string.Format(@"Форма обучения: {0}
Условия обучения: {1}", info.StudyFormName, basis);

                    Paragraph p = new Paragraph(header, font);
                    document.Add(p);

                    float midStr = 13f;
                    p = new Paragraph(20f);
                    p.Add(new Phrase("ПРОТОКОЛ № ", new Font(bfTimes, 14, Font.BOLD)));
                    p.Add(new Phrase(info.Number, new Font(bfTimes, 18, Font.BOLD)));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    p = new Paragraph(midStr);
                    p.Add(new Phrase(@"заседания Приемной комиссии Санкт-Петербургского Государственного Университета
об изменении типа конкурса целевикам ", new Font(bfTimes, 10, Font.BOLD)));

                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    //date
                    p = new Paragraph(midStr);
                    p.Add(new Paragraph(string.Format("от {0}", Util.GetDateString(info.Date, true, true)), new Font(bfTimes, 10, Font.BOLD)));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    string spec = "";
                    PdfPTable curT = null;
                    int cnt = 0;
                    string currSpec = null;

                    var lst = ProtocolDataProvider.GetProtocolData(gProtocolId);
                    foreach (var v in lst)
                    {
                        cnt++;

                        currSpec = v.Direction;
                        if (spec != currSpec)
                        {
                            spec = currSpec;
                            cnt = 1;

                            if (curT != null)
                                document.Add(curT);

                            //Table
                            Table table = new Table(6);
                            table.Padding = 3;
                            table.Spacing = 0;
                            float[] headerwidths = { 5, 10, 30, 15, 10, 10 };
                            table.Widths = headerwidths;
                            table.Width = 100;

                            PdfPTable t = new PdfPTable(6);
                            t.SetWidthPercentage(headerwidths, document.PageSize);
                            t.WidthPercentage = 100f;
                            t.SpacingBefore = 10f;
                            t.SpacingAfter = 10f;

                            t.HeaderRows = 2;

                            Phrase pra = new Phrase(string.Format("По направлению {0} ", currSpec), new Font(bfTimes, 10));

                            PdfPCell pcell = new PdfPCell(pra);
                            pcell.BorderWidth = 0;
                            pcell.Colspan = 7;
                            t.AddCell(pcell);

                            string[] headers = new string[]
                                {
                                    "№ п/п",
                                    "Рег.номер",
                                    "ФАМИЛИЯ, ИМЯ, ОТЧЕСТВО",
                                    "Номер аттестата или диплома",                            
                                    "Новый вид конкурса",
                                    "Примечания"
                                };
                            foreach (string h in headers)
                            {
                                PdfPCell cell = new PdfPCell();
                                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                cell.AddElement(new Phrase(h, new Font(bfTimes, 10, Font.BOLD)));

                                t.AddCell(cell);
                            }

                            curT = t;
                        }

                        curT.AddCell(new Phrase(cnt.ToString(), new Font(bfTimes, 10)));
                        curT.AddCell(new Phrase(v.RegNum, new Font(bfTimes, 10)));
                        curT.AddCell(new Phrase(v.FIO, new Font(bfTimes, 10)));
                        curT.AddCell(new Phrase(v.EducationDocument, new Font(bfTimes, 10)));
                        curT.AddCell(new Phrase(v.CompetitionName, new Font(bfTimes, 10)));
                        curT.AddCell(new Phrase(v.Comment, new Font(bfTimes, 10)));
                    }


                    if (curT != null)
                        document.Add(curT);

                    //FOOTER
                    p = new Paragraph(30f);
                    p.KeepTogether = true;
                    p.Add(new Phrase("Ответственный секретарь Приемной комиссии СПбГУ____________________________________________________________", new Font(bfTimes, 10)));
                    document.Add(p);

                    p = new Paragraph();
                    p.Add(new Phrase("Заместитель начальника Управления по организации приема – советник проректора по направлениям___________________", new Font(bfTimes, 10)));
                    document.Add(p);

                    p = new Paragraph();
                    p.Add(new Phrase("Ответственный секретарь комиссии по приему документов_______________________________________________________", new Font(bfTimes, 10)));
                    document.Add(p);

                    document.Close();

                    Process pr = new Process();
                    if (forPrint)
                    {
                        pr.StartInfo.Verb = "Print";
                        pr.StartInfo.FileName = string.Format(savePath);
                        pr.Start();
                    }
                    else
                    {
                        pr.StartInfo.Verb = "Open";
                        pr.StartInfo.FileName = string.Format(savePath);
                        pr.Start();
                    }
                }

            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
            finally
            {
                if (fileS != null)
                    fileS.Dispose();
            }
        }
        public static void PrintChangeCompBEProtocol(string protocolId, bool forPrint, string savePath)
        {
            FileStream fileS = null;
            try
            {
                Guid gProtocolId = Guid.Parse(protocolId);
                var info = ProtocolDataProvider.GetProtocolInfo(gProtocolId, 6);

                string basis = string.Empty;
                switch (info.StudyBasisId.ToString())
                {
                    case "1":
                        basis = "Бюджетные места";
                        break;
                    case "2":
                        basis = "Места по договорам с оплатой стоимости обучения";
                        break;
                }

                Document document = new Document(PageSize.A4.Rotate(), 50, 50, 50, 50);
                using (fileS = new FileStream(savePath, FileMode.Create))
                {
                    BaseFont bfTimes = BaseFont.CreateFont(string.Format(@"{0}\times.ttf", MainClass.dirTemplates), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    Font font = new Font(bfTimes, 10);

                    PdfWriter.GetInstance(document, fileS);
                    document.Open();

                    //HEADER
                    string header = string.Format(@"Форма обучения: {0}
Условия обучения: {1}", info.StudyFormName, basis);

                    Paragraph p = new Paragraph(header, font);
                    document.Add(p);

                    float midStr = 13f;
                    p = new Paragraph(20f);
                    p.Add(new Phrase("ПРОТОКОЛ № ", new Font(bfTimes, 14, Font.BOLD)));
                    p.Add(new Phrase(info.Number, new Font(bfTimes, 18, Font.BOLD)));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    p = new Paragraph(midStr);
                    p.Add(new Phrase(@"заседания Приемной комиссии Санкт-Петербургского Государственного Университета
об изменении типа конкурса на общий ", new Font(bfTimes, 10, Font.BOLD)));

                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    //date
                    p = new Paragraph(midStr);
                    p.Add(new Paragraph(string.Format("от {0}", Util.GetDateString(info.Date, true, true)), new Font(bfTimes, 10, Font.BOLD)));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    string spec = "";
                    PdfPTable curT = null;
                    int cnt = 0;
                    string currSpec = null;

                    var lst = ProtocolDataProvider.GetProtocolData(gProtocolId);
                    foreach (var v in lst)
                    {
                        cnt++;

                        currSpec = v.Direction;
                        if (spec != currSpec)
                        {
                            spec = currSpec;
                            cnt = 1;

                            if (curT != null)
                                document.Add(curT);

                            //Table
                            Table table = new Table(6);
                            table.Padding = 3;
                            table.Spacing = 0;
                            float[] headerwidths = { 5, 10, 30, 15, 10, 10 };
                            table.Widths = headerwidths;
                            table.Width = 100;

                            PdfPTable t = new PdfPTable(6);
                            t.SetWidthPercentage(headerwidths, document.PageSize);
                            t.WidthPercentage = 100f;
                            t.SpacingBefore = 10f;
                            t.SpacingAfter = 10f;

                            t.HeaderRows = 2;

                            Phrase pra = new Phrase(string.Format("По направлению {0} ", currSpec), new Font(bfTimes, 10));
                            PdfPCell pcell = new PdfPCell(pra);
                            pcell.BorderWidth = 0;
                            pcell.Colspan = 7;
                            t.AddCell(pcell);

                            string[] headers = new string[]
                                {
                                    "№ п/п",
                                    "Рег.номер",
                                    "ФАМИЛИЯ, ИМЯ, ОТЧЕСТВО",
                                    "Номер аттестата или диплома",                            
                                    "Новый вид конкурса",
                                    "Примечания"
                                };

                            foreach (string h in headers)
                            {
                                PdfPCell cell = new PdfPCell();
                                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                cell.AddElement(new Phrase(h, new Font(bfTimes, 10, Font.BOLD)));

                                t.AddCell(cell);
                            }

                            curT = t;
                        }

                        curT.AddCell(new Phrase(cnt.ToString(), new Font(bfTimes, 10)));
                        curT.AddCell(new Phrase(v.RegNum, new Font(bfTimes, 10)));
                        curT.AddCell(new Phrase(v.FIO, new Font(bfTimes, 10)));
                        curT.AddCell(new Phrase(v.EducationDocument, new Font(bfTimes, 10)));
                        curT.AddCell(new Phrase(v.CompetitionName, new Font(bfTimes, 10)));
                        curT.AddCell(new Phrase(v.Comment, new Font(bfTimes, 10)));
                    }


                    if (curT != null)
                        document.Add(curT);

                    //FOOTER
                    p = new Paragraph(30f);
                    p.KeepTogether = true;
                    p.Add(new Phrase("Ответственный секретарь Приемной комиссии СПбГУ____________________________________________________________", new Font(bfTimes, 10)));
                    document.Add(p);

                    p = new Paragraph();
                    p.Add(new Phrase("Заместитель начальника Управления по организации приема – советник проректора по направлениям___________________", new Font(bfTimes, 10)));
                    document.Add(p);

                    p = new Paragraph();
                    p.Add(new Phrase("Ответственный секретарь комиссии по приему документов_______________________________________________________", new Font(bfTimes, 10)));
                    document.Add(p);

                    document.Close();

                    Process pr = new Process();
                    if (forPrint)
                    {
                        pr.StartInfo.Verb = "Print";
                        pr.StartInfo.FileName = string.Format(savePath);
                        pr.Start();
                    }
                    else
                    {
                        pr.StartInfo.Verb = "Open";
                        pr.StartInfo.FileName = string.Format(savePath);
                        pr.Start();
                    }
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
            finally
            {
                if (fileS != null)
                    fileS.Dispose();
            }
        }
        public static void PrintEntryView(string sProtocolId, string savePath)
        {
            FileStream fileS = null;
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    Guid gProtocolId = new Guid(sProtocolId);
                    var ProtocolInfo = ProtocolDataProvider.GetProtocolInfo(gProtocolId, 4);

                    string docNum = ProtocolInfo.Number.ToString();
                    DateTime docDate = ProtocolInfo.Date.Date;

                    var SF = context.StudyForm.Where(x => x.Id == ProtocolInfo.StudyFormId).FirstOrDefault();
                    string form = SF.Acronym;
                    string form2 = SF.RodName;
                    string facDat = ProtocolInfo.FacultyDatName;

                    string basis = string.Empty;
                    switch (ProtocolInfo.StudyBasisId)
                    {
                        case 1:
                            basis = "обучение за счет средств федерального бюджета";
                            break;
                        case 2:
                            basis = "обучение по договорам с оплатой стоимости обучения";
                            break;
                    }

                    string list = string.Empty, sec = string.Empty;

                    string copyDoc = "оригиналы";
                    if (ProtocolInfo.IsListener)
                    {
                        list = " в качестве слушателя";
                        copyDoc = "заверенные ксерокопии";
                    }
                    if (ProtocolInfo.IsReduced)
                        sec = " (сокращенной)";
                    if (ProtocolInfo.IsParallel)
                        sec = " (параллельной)";
                    if (ProtocolInfo.IsSecond)
                        sec = " (сокращенной)";

                    Document document = new Document(PageSize.A4, 50, 50, 50, 50);
                    using (fileS = new FileStream(savePath, FileMode.Create))
                    {
                        BaseFont bfTimes = BaseFont.CreateFont(string.Format(@"{0}\times.ttf", MainClass.dirTemplates), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                        Font font = new Font(bfTimes, 12);

                        PdfWriter writer = PdfWriter.GetInstance(document, fileS);
                        document.Open();

                        float firstLineIndent = 30f;
                        //HEADER
                        Paragraph p = new Paragraph("Правительство Российской Федерации", new Font(bfTimes, 12, Font.BOLD));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        p = new Paragraph("Федеральное государственное бюджетное образовательное учреждение", new Font(bfTimes, 12));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        p = new Paragraph("высшего профессионального образования", new Font(bfTimes, 12));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        p = new Paragraph("САНКТ-ПЕТЕРБУРГСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ", new Font(bfTimes, 12, Font.BOLD));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        p = new Paragraph("ПРЕДСТАВЛЕНИЕ", new Font(bfTimes, 20, Font.BOLD));
                        p.Alignment = Element.ALIGN_CENTER;
                        document.Add(p);

                        p = new Paragraph(string.Format("От {0} г. № {1}", Util.GetDateString(docDate, true, true), docNum), font);
                        p.SpacingBefore = 10f;
                        document.Add(p);

                        p = new Paragraph(10f);
                        p.Add(new Paragraph("по " + facDat, font));

                        string naprspecRod = "",
                            profspec = "",
                            naprobProgRod = "",
                            educDoc = ""; ;

                        naprobProgRod = "образовательной программе";
                        naprspecRod = "направлению";

                        if (MainClass.dbType == PriemType.PriemMag)
                        {
                            profspec = "профилю";
                            educDoc = "о высшем профессиональном образовании";
                        }
                        else
                        {
                            profspec = "профилю";
                            educDoc = "об образовании";
                        }

                        p.Add(new Paragraph(string.Format("по основной{4} образовательной программе подготовки {0} на направление {1} «{2}» ", ProtocolInfo.StudyLevelNameRod, ProtocolInfo.LicenseProgramCode, ProtocolInfo.LicenseProgramName, sec), font));
                        p.Add(new Paragraph((form + " форма обучения,").ToLower(), font));
                        p.Add(new Paragraph(basis, font));
                        p.IndentationLeft = 320;
                        document.Add(p);

                        p = new Paragraph();
                        p.Add(new Paragraph("О зачислении на 1 курс", font));
                        p.SpacingBefore = 10f;
                        document.Add(p);

                        p = new Paragraph(string.Format("В соответствии с Федеральным законом  от 22.08.1996 N 125-ФЗ \"О высшем и послевузовском профессиональном образовании\", Порядком приема граждан в образовательные учреждения высшего профессионального образования, утвержденным Приказом Минобрнауки РФ от 28.12.2011 N 2895, Правилами приема в Санкт-Петербургский государственный университет на основные образовательные программы высшего профессионального образования (программы бакалавриата, программы подготовки специалиста, программы магистратуры) в {0} году", MainClass.iPriemYear), font);
                        p.SpacingBefore = 10f;
                        p.FirstLineIndent = firstLineIndent;
                        p.Alignment = Element.ALIGN_JUSTIFIED;
                        document.Add(p);

                        p = new Paragraph(string.Format("Представить на рассмотрение Приемной комиссии СПбГУ по вопросу зачисления c 01.09.{4} года на 1 курс{2} с освоением основной{3} образовательной программы подготовки {0} по {1} форме обучения следующих граждан, успешно выдержавших вступительные испытания:", ProtocolInfo.StudyLevelNameRod, form2, list, sec, MainClass.iPriemYear), font);
                        p.SpacingBefore = 20f;
                        p.FirstLineIndent = firstLineIndent;
                        p.Alignment = Element.ALIGN_JUSTIFIED;
                        
                        document.Add(p);

                        string curSpez = "-";
                        string curObProg = "-";
                        string curHeader = "-";

                        int counter = 0;

                        var lst = ProtocolDataProvider.GetEntryViewData(gProtocolId, false);
                        foreach (var v in lst)
                        {
                            ++counter;
                            string obProg = v.ObrazProgram;
                            string obProgCrypt = v.ObrazProgramCrypt;
                            string obProgId = v.ObrazProgramId.ToString();

                            if (obProgId != curObProg)
                            {
                                p = new Paragraph();
                                p.Add(new Paragraph(string.Format("{3}по {0} {1} \"{2}\"", naprspecRod, v.LicenseProgramCode, v.LicenseProgramName, curObProg == "-" ? "" : "\r\n"), font));

                                if (!string.IsNullOrEmpty(obProg))
                                    p.Add(new Paragraph(string.Format("по {0} {1} \"{2}\"", naprobProgRod, obProgCrypt, obProg), font));

                                string spez = v.ProfileName;
                                if (spez != curSpez)
                                {
                                    if (!string.IsNullOrEmpty(spez) && spez != "нет")
                                        p.Add(new Paragraph(string.Format("по {0} \"{1}\"", profspec, spez), font));

                                    curSpez = spez;
                                }

                                p.IndentationLeft = 40;
                                document.Add(p);

                                curObProg = obProgId;
                            }
                            else
                            {
                                string spez = v.ProfileName;
                                if (spez != curSpez && spez != "нет")
                                {
                                    p = new Paragraph();
                                    p.Add(new Paragraph(string.Format("{3}по {0} {1} \"{2}\"", naprspecRod, v.LicenseProgramCode, v.LicenseProgramName, curObProg == "-" ? "" : "\r\n"), font));

                                    if (!string.IsNullOrEmpty(obProg))
                                        p.Add(new Paragraph(string.Format("по {0} \"{1}\"", naprobProgRod, obProg), font));

                                    if (!string.IsNullOrEmpty(spez))
                                        p.Add(new Paragraph(string.Format("по {0} \"{1}\"", profspec, spez), font));

                                    p.IndentationLeft = 40;
                                    document.Add(p);

                                    curSpez = spez;
                                }
                            }

                            string header = v.EntryHeaderName;
                            if (header != curHeader)
                            {
                                p = new Paragraph();
                                p.Add(new Paragraph(string.Format("\r\n{0}:", header), font));
                                p.IndentationLeft = 40;
                                document.Add(p);

                                curHeader = header;
                            }

                            p = new Paragraph();
                            p.Add(new Paragraph(string.Format("{0}. {1} {2}", counter, v.FIO, v.TotalSum.ToString()), font));
                            p.IndentationLeft = 60;
                            document.Add(p);
                        }

                        //FOOTER
                        p = new Paragraph();
                        p.SpacingBefore = 30f;
                        p.Alignment = Element.ALIGN_JUSTIFIED;
                        p.FirstLineIndent = firstLineIndent;
                        p.Add(new Phrase("ОСНОВАНИЕ:", new Font(bfTimes, 12)));
                        p.Add(new Phrase(string.Format(" личные заявления, протоколы вступительных испытаний, {0} документов государственного образца {1}.", copyDoc, educDoc), font));
                        document.Add(p);

                        p = new Paragraph();
                        p.SpacingBefore = 30f;
                        p.KeepTogether = true;
                        p.Add(new Paragraph("Ответственный секретарь", font));
                        p.Add(new Paragraph("комиссии по приему документов СПбГУ                                                                                          ", font));
                        document.Add(p);

                        p = new Paragraph();
                        p.SpacingBefore = 30f;
                        p.Add(new Paragraph("Заместитель начальника управления - ", font));
                        p.Add(new Paragraph("советник проректора по направлениям", font));
                        document.Add(p);

                        document.Close();

                        Process pr = new Process();

                        pr.StartInfo.Verb = "Open";
                        pr.StartInfo.FileName = string.Format(savePath);
                        pr.Start();
                    }
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
            finally
            {
                if (fileS != null)
                    fileS.Dispose();
            }
        }

        public static void PrintOrder(Guid gProtocolId, bool isRus, bool isCel)
        {
            try
            {
                WordDoc wd = new WordDoc(string.Format(@"{0}\EntryOrder.dot", MainClass.dirTemplates));
                TableDoc td = wd.Tables[0];

                var ProtocolInfo = ProtocolDataProvider.GetProtocolInfo(gProtocolId, 4);

                string docNum;
                DateTime? docDate;

                string basis = string.Empty;
                string basis2 = string.Empty;
                string form = string.Empty;
                string form2 = string.Empty;
                string sLicenseProgramName = string.Empty;
                string sLicenseProgramCode = string.Empty;
                string sStudyLevelNameRod = string.Empty;
                using (PriemEntities ctx = new PriemEntities())
                {
                    docNum = (from protocol in ctx.OrderNumbers
                              where protocol.ProtocolId == gProtocolId
                              select protocol.ComissionNumber).DefaultIfEmpty("НЕ УКАЗАН").FirstOrDefault();

                    docDate = (from protocol in ctx.OrderNumbers
                               where protocol.ProtocolId == gProtocolId
                               select protocol.ComissionDate).FirstOrDefault();

                    sLicenseProgramName =
                        (from entry in ctx.extEntry
                         join extentryView in ctx.extEntryView on entry.LicenseProgramId equals extentryView.LicenseProgramId
                         where extentryView.Id == gProtocolId
                         select entry.LicenseProgramName).FirstOrDefault();

                    sLicenseProgramCode =
                        (from entry in ctx.extEntry
                         join extentryView in ctx.extEntryView on entry.LicenseProgramId equals extentryView.LicenseProgramId
                         where extentryView.Id == gProtocolId
                         select entry.LicenseProgramCode).FirstOrDefault();

                    sStudyLevelNameRod =
                        (from entry in ctx.Entry
                         join extentryView in ctx.extEntryView on entry.LicenseProgramId equals extentryView.LicenseProgramId
                         where extentryView.Id == gProtocolId
                         select entry.StudyLevel.NameRod).FirstOrDefault();

                    var SF = ctx.StudyForm.Where(x => x.Id == ProtocolInfo.StudyFormId).Select(x => new { x.Name, x.RodName }).FirstOrDefault();
                    form = SF.Name + " форма обучения";
                    form2 = "по " + SF.RodName + " форме";
                }

                string educDoc = "", list = "", sec = "";

                if (ProtocolInfo.IsListener)
                    list = " в качестве слушателя";
                if (ProtocolInfo.IsSecond)
                    sec += " (для лиц с ВО)";
                if (ProtocolInfo.IsParallel)
                    sec += " (параллельное обучение)";
                if (ProtocolInfo.IsReduced)
                    sec += " (сокращенной)";

                string dogovorDoc = "";
                switch (ProtocolInfo.StudyBasisId)
                {
                    case 1:
                        basis2 = "обучения за счет бюджетных ассигнований федерального бюджета";
                        dogovorDoc = "";
                        educDoc = ", оригиналы документа установленного образца об образовании";
                        break;
                    case 2:
                        basis2 = "обучения по договорам об образовании";
                        dogovorDoc = ", договоры об образовании";
                        educDoc = "";
                        break;
                }

                wd.SetFields("Граждан", isRus ? "граждан Российской Федерации" : "иностранных граждан");
                wd.SetFields("Граждан2", isRus ? "граждан Российской Федерации" : "");
                wd.SetFields("Стипендия", (ProtocolInfo.StudyBasisId == 2 || ProtocolInfo.StudyFormId == 2) ? "" : "и назначении стипендии");
                wd.SetFields("Форма2", form2);
                wd.SetFields("Основа2", basis2);
                wd.SetFields("БакСпецРод", sStudyLevelNameRod);
                wd.SetFields("Слушатель", list);
                wd.SetFields("Сокращ", sec);

                wd.SetFields("ДатаПриказа", docDate.HasValue ? docDate.Value.ToShortDateString() : "НЕТ ДАТЫ");
                wd.SetFields("НомерПриказа", docNum);

                wd.SetFields("DogovorDoc", dogovorDoc);
                wd.SetFields("EducDoc", educDoc);

                int curRow = 4, counter = 0;
                string curProfileName = "нет";
                string curObrazProgramId = "-";
                string curHeader = "-";
                string curCountry = "-";
                string curLPHeader = "-";
                string curMotivation = "-";
                string Motivation = string.Empty;


                var lst = ProtocolDataProvider.GetEntryViewData(gProtocolId, isRus);
                bool bFirstRun = true;
                foreach (var v in lst)
                {
                    ++counter;

                    string header = v.EntryHeaderName;

                    if (!isCel && !bFirstRun)
                    {
                        if (header != curHeader)
                        {
                            AddRowInTableOrder(header, ref td, ref curRow);
                            curHeader = header;
                        }
                    }

                    bFirstRun = false;

                    string LP = v.LicenseProgramName;
                    if (curLPHeader != LP)
                    {
                        AddRowInTableOrder(string.Format("{2}\tпо направлению подготовки {0} \"{1}\"", v.LicenseProgramCode, LP, curObrazProgramId == "-" ? "" : "\r\n"), ref td, ref curRow);
                        curLPHeader = LP;
                    }

                    string ObrazProgramId = v.ObrazProgramId.ToString();
                    if (ObrazProgramId != curObrazProgramId)
                    {
                        if (!string.IsNullOrEmpty(v.ObrazProgram))
                            AddRowInTableOrder(string.Format("\tпо образовательной программе {0} \"{1}\"", v.ObrazProgramCrypt, v.ObrazProgram), ref td, ref curRow);

                        string profileName = v.ProfileName;
                        if (!string.IsNullOrEmpty(profileName) && profileName != "нет")
                            AddRowInTableOrder(string.Format("\tпо профилю \"{0}\"", profileName), ref td, ref curRow);

                        curProfileName = profileName;
                        curObrazProgramId = ObrazProgramId;

                        if (!isCel)
                        {
                            if (header != curHeader)
                            {
                                AddRowInTableOrder(string.Format("\t{0}:", header), ref td, ref curRow);
                                curHeader = header;
                            }
                        }
                    }
                    else
                    {
                        string profileName = v.ProfileName;
                        if (profileName != curProfileName)
                        {
                            if (!string.IsNullOrEmpty(profileName) && profileName != "нет")
                                AddRowInTableOrder(string.Format("\tпо профилю \"{0}\"", profileName), ref td, ref curRow);

                            curProfileName = profileName;
                            if (!isCel)
                                AddRowInTableOrder(string.Format("\t{0}:", header), ref td, ref curRow);
                        }
                    }

                    if (!isRus)
                    {
                        string country = v.CountryNameRod;
                        if (country != curCountry)
                        {
                            AddRowInTableOrder(string.Format("\r\n граждан {0}:", country), ref td, ref curRow);
                            curCountry = country;
                        }
                    }

                    string balls = v.TotalSum.HasValue ? v.TotalSum.Value.ToString() : "";
                    string ballToStr = GetBallsToStr(balls);

                    if (isCel && curMotivation == "-")
                        curMotivation = string.Format("ОСНОВАНИЕ: договор об организации целевого приема с {0} от … № …, Протокол заседания Приемной комиссии СПбГУ от ДАТА № ..., личное заявление, оригинал документа государственного образца об образовании.", v.CelCompName);
                    string tmpMotiv = curMotivation;
                    Motivation = string.Format("ОСНОВАНИЕ: договор об организации целевого приема с {0} от … № …, Протокол заседания Приемной комиссии СПбГУ от ДАТА № ..., личное заявление, оригинал документа государственного образца об образовании.", v.CelCompName);

                    if (isCel && curMotivation != Motivation)
                    {
                        string CelCompText = v.CelCompName;
                        Motivation = string.Format("ОСНОВАНИЕ: договор об организации целевого приема с {0} от … № …, Протокол заседания Приемной комиссии СПбГУ от ДАТА № .., личное заявление, оригинал документа государственного образца об образовании.", CelCompText);
                        curMotivation = Motivation;
                    }
                    else
                        Motivation = string.Empty;

                    AddRowInTableOrder(string.Format("\t\t1.{0}. {1} {2} {3}", counter, v.FIO, balls + ballToStr, string.IsNullOrEmpty(Motivation) ? "" : ("\n\n\t\t" + tmpMotiv + "\n")), ref td, ref curRow);
                }

                if (!string.IsNullOrEmpty(curMotivation) && isCel)
                    td[0, curRow] += "\n\t\t" + curMotivation + "\n";

                //платникам и всем очно-заочникам стипендия не платится
                if (ProtocolInfo.StudyBasisId != 2 && ProtocolInfo.StudyFormId != 2)
                    AddRowInTableOrder("\r\n2.    Назначить лицам, указанным в п. 1 настоящего приказа, стипендию в размере 1340 рублей ежемесячно с 01.09.2014 по 31.01.2015.", ref td, ref curRow);
            }
            catch (WordException we)
            {
                WinFormsServ.Error(we);
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc);
            }
        }
        public static void PrintOrderReview(Guid gProtocolId, bool isRus)
        {
            try
            {
                WordDoc wd = new WordDoc(string.Format(@"{0}\EntryOrderList.dot", MainClass.dirTemplates));

                var ProtocolInfo = ProtocolDataProvider.GetProtocolInfo(gProtocolId, 4);
                using (PriemEntities ctx = new PriemEntities())
                {
                    string sLicenseProgramName =
                        (from entry in ctx.extEntry
                         join extentryView in ctx.extEntryView on entry.LicenseProgramId equals extentryView.LicenseProgramId
                         where extentryView.Id == gProtocolId
                         select entry.LicenseProgramName).FirstOrDefault();

                    string sLicenseProgramCode =
                        (from entry in ctx.extEntry
                         join extentryView in ctx.extEntryView on entry.LicenseProgramId equals extentryView.LicenseProgramId
                         where extentryView.Id == gProtocolId
                         select entry.LicenseProgramCode).FirstOrDefault();

                    string sStudyLevelNameRod =
                        (from entry in ctx.Entry
                         join extentryView in ctx.extEntryView on entry.LicenseProgramId equals extentryView.LicenseProgramId
                         where extentryView.Id == gProtocolId
                         select entry.StudyLevel.NameRod).FirstOrDefault();

                    string basis = string.Empty, educDoc = string.Empty;
                    switch (ProtocolInfo.StudyBasisId)
                    {
                        case 1:
                            basis = "за счет бюджетных ассигнований федерального бюджета";
                            educDoc = ", оригиналы документа установленного образца об образовании";
                            break;
                        case 2:
                            basis = "по договорам об образовании";
                            educDoc = ", договоры об образовании";
                            break;
                    }

                    var SF = ctx.StudyForm.Where(x => x.Id == ProtocolInfo.StudyFormId).Select(x => new { x.Name, x.RodName }).FirstOrDefault();
                    string form2 = "по " + SF.RodName + " форме";

                    int curRow = 5, counter = 0;
                    TableDoc td = null;

                    DateTime? dtComissionDate =
                        (from protocol in ctx.OrderNumbers
                         where protocol.ProtocolId == gProtocolId
                         select protocol.ComissionDate).FirstOrDefault();

                    string sComissionNum =
                        (from protocol in ctx.OrderNumbers
                         where protocol.ProtocolId == gProtocolId
                         select protocol.ComissionNumber).DefaultIfEmpty("НЕ УКАЗАН").FirstOrDefault();

                    string docNum =
                        (from orderNumbers in ctx.OrderNumbers
                         where orderNumbers.ProtocolId == gProtocolId
                         select isRus ? orderNumbers.OrderNum : orderNumbers.OrderNumFor).FirstOrDefault();
                    if (string.IsNullOrEmpty(docNum))
                        docNum = "НЕТ НОМЕРА";

                    DateTime? tempDate =
                        (from orderNumbers in ctx.OrderNumbers
                         where orderNumbers.ProtocolId == gProtocolId
                         select isRus ? orderNumbers.OrderDate : orderNumbers.OrderDateFor).FirstOrDefault();
                    
                    string docDate = tempDate.HasValue ? tempDate.Value.ToShortDateString() : "НЕТ ДАТЫ";

                    var lst = ProtocolDataProvider.GetEntryViewData(gProtocolId, isRus);
                    foreach (var v in lst)
                    {
                        if (v.CompetitionId == 11 || v.CompetitionId == 12)
                            wd.InsertAutoTextInEnd("выпискаКРЫМ", true);
                        else
                            wd.InsertAutoTextInEnd("выписка", true);

                        wd.GetLastFields(13);
                        td = wd.Tables[counter];

                        wd.SetFields("Граждан", isRus ? "граждан РФ" : "иностранных граждан");
                        wd.SetFields("Граждан2", isRus ? "граждан Российской Федерации" : "");
                        wd.SetFields("Стипендия", (ProtocolInfo.StudyBasisId == 2 || ProtocolInfo.StudyFormId == 2) ? "" : "и назначении стипендии");
                        wd.SetFields("Форма2", form2);
                        wd.SetFields("Основа2", basis);
                        wd.SetFields("БакСпецРод", sStudyLevelNameRod);
                        wd.SetFields("ПриказДата", docDate);
                        wd.SetFields("ПриказНомер", "№ " + docNum);
                        wd.SetFields("SignerName", v.SignerName);
                        wd.SetFields("SignerPosition", v.SignerPosition);
                        wd.SetFields("Основание", educDoc);
                        if (dtComissionDate.HasValue)
                            wd.SetFields("ДатаОснования", ((DateTime)dtComissionDate).ToShortDateString());
                        else
                            wd.SetFields("ДатаОснования", "ДАТА");
                        wd.SetFields("НомерОснования", sComissionNum ?? "НОМЕР");

                        string curLPHeader = "-";
                        string curSpez = "-";
                        string curObProg = "-";
                        string curHeader = "-";
                        string curCountry = "-";

                        ++counter;

                        string LP = v.LicenseProgramCode + " " + v.LicenseProgramName;
                        if (curLPHeader != LP)
                        {
                            AddRowInTableOrder(string.Format("{1}\tпо направлению подготовки \"{0}\"", LP, curObProg == "-" ? "" : "\r\n"), ref td, ref curRow);
                            curLPHeader = LP;
                        }

                        string obProg = v.ObrazProgram;
                        if (obProg != curObProg)
                        {
                            if (!string.IsNullOrEmpty(obProg))
                                AddRowInTableOrder(string.Format("\tпо образовательной программе {0} \"{1}\"", v.ObrazProgramCrypt, obProg), ref td, ref curRow);

                            string spez = v.ProfileName;
                            if (!string.IsNullOrEmpty(spez) && spez != "нет")
                                AddRowInTableOrder(string.Format("\t профилю \"{0}\"", spez), ref td, ref curRow);

                            curSpez = spez;
                            curObProg = obProg;
                        }
                        else
                        {
                            string spez = v.ProfileName;
                            if (spez != curSpez)
                            {
                                if (!string.IsNullOrEmpty(spez) && spez != "нет")
                                    AddRowInTableOrder(string.Format("\t профилю \"{0}\"", spez), ref td, ref curRow);

                                curSpez = spez;
                            }
                        }

                        if (!isRus)
                        {
                            string country = v.CountryNameRod;
                            if (country != curCountry)
                            {
                                AddRowInTableOrder(string.Format("\r\n граждан {0}:", country), ref td, ref curRow);
                                curCountry = country;
                            }
                        }

                        string header = v.EntryHeaderName;
                        if (header != curHeader)
                        {
                            AddRowInTableOrder(string.Format("\t{0}:", header), ref td, ref curRow);
                            curHeader = header;
                        }

                        string balls = v.TotalSum.ToString();
                        AddRowInTableOrder(string.Format("\t\t{0} {1}", v.FIO, balls + GetBallsToStr(balls)), ref td, ref curRow);

                        if (ProtocolInfo.StudyBasisId != 2 && ProtocolInfo.StudyFormId != 2)
                            AddRowInTableOrder("\r\n2.      Назначить указанным лицам стипендию в размере 1340 рубля ежемесячно до 31 января 2015 г.", ref td, ref curRow);
                    }
                }
            }
            catch (WordException we)
            {
                WinFormsServ.Error(we.Message);
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
        }
        private static void AddRowInTableOrder(string text, ref TableDoc td, ref int curRow)
        {
            td.AddRow(1);
            curRow++;
            td[0, curRow] = text;
        }

        private static string GetBallsToStr(string balls)
        {
            string ballToStr = " балл";

            if (balls.Length == 0)
                ballToStr = "";
            else if (balls.EndsWith("1"))
            {
                if (balls.EndsWith("11"))
                    ballToStr += "ов";
                else
                    ballToStr += "";
            }
            else if (balls.EndsWith("2") || balls.EndsWith("3") || balls.EndsWith("4"))
            {
                if (balls.EndsWith("12") || balls.EndsWith("13") || balls.EndsWith("14"))
                    ballToStr += "ов";
                else
                    ballToStr += "а";
            }
            else
                ballToStr += "ов";

            return ballToStr;
        }

        public static void PrintDisEntryOrder(string protocolId, bool isRus)
        {
            try
            {
                WordDoc wd = new WordDoc(string.Format(@"{0}\DisEntryOrder.dot", MainClass.dirTemplates));
                TableDoc td = wd.Tables[0];

                Guid gProtocolId = Guid.Parse(protocolId);
                var ProtocolInfo = ProtocolDataProvider.GetProtocolInfo(gProtocolId, 5);

                using (PriemEntities ctx = new PriemEntities())
                {
                    Guid entryProtocolId =
                        (from extEntryView in ctx.extEntryView_ForDisEntered
                         join extDisEntryView in ctx.extDisEntryView on extEntryView.AbiturientId equals extDisEntryView.AbiturientId
                         where !extDisEntryView.IsOld && extDisEntryView.Id == gProtocolId
                         select extEntryView.Id).FirstOrDefault();

                    string docNum = "НОМЕР";
                    string docDate = "ДАТА";
                    
                    DateTime? tempDate;
                    docNum =
                        (from orderNumbers in ctx.OrderNumbers
                         where orderNumbers.ProtocolId == entryProtocolId
                         select isRus ? orderNumbers.OrderNum : orderNumbers.OrderNumFor).FirstOrDefault();

                    tempDate = (DateTime?)
                        (from orderNumbers in ctx.OrderNumbers
                         where orderNumbers.ProtocolId == entryProtocolId
                         select isRus ? orderNumbers.OrderDate : orderNumbers.OrderDateFor).FirstOrDefault();

                    if (tempDate.HasValue)
                        docDate = tempDate.Value.ToShortDateString();
                    else
                        docDate = "!НЕТ ДАТЫ";

                    string facDat =
                        (from protocol in ctx.Protocol
                         join sP_Faculty in ctx.SP_Faculty on protocol.FacultyId equals sP_Faculty.Id
                         where protocol.Id == gProtocolId
                         select sP_Faculty.DatName).FirstOrDefault();

                    string list = string.Empty, sec = string.Empty;
                    if (ProtocolInfo.IsSecond)
                        list = " в качестве слушателя";
                    if (ProtocolInfo.IsReduced)
                        sec += " (сокращенной)";
                    if (ProtocolInfo.IsListener)
                        sec += " (для лиц с высшим образованием)";

                    string LicenseProgramName =
                        (from entry in ctx.extEntry
                         join extdisEntryView in ctx.extDisEntryView on entry.LicenseProgramId equals extdisEntryView.LicenseProgramId
                         where extdisEntryView.Id == gProtocolId && extdisEntryView.StudyLevelGroupId == MainClass.studyLevelGroupId
                         select entry.LicenseProgramName).FirstOrDefault();

                    string LicenseProgramCode =
                        (from entry in ctx.extEntry
                         join extdisEntryView in ctx.extDisEntryView on entry.LicenseProgramId equals extdisEntryView.LicenseProgramId
                         where extdisEntryView.Id == gProtocolId && extdisEntryView.StudyLevelGroupId == MainClass.studyLevelGroupId
                         select entry.LicenseProgramCode).FirstOrDefault();

                    string StudyLevelName =
                        (from entry in ctx.extEntry
                         join extdisEntryView in ctx.extDisEntryView on entry.LicenseProgramId equals extdisEntryView.LicenseProgramId
                         where extdisEntryView.Id == gProtocolId && extdisEntryView.StudyLevelGroupId == MainClass.studyLevelGroupId
                         select entry.StudyLevelName).FirstOrDefault();

                    string basis = string.Empty;
                    switch (ProtocolInfo.StudyBasisId)
                    {
                        case 1:
                            basis = "обучение за счет средств федерального бюджета";
                            break;
                        case 2:
                            basis = string.Format("по договорам оказания государственной услуги по обучению по основной{0} образовательной программе высшего профессионального образования", sec);
                            break;
                    }

                    var SF = ctx.StudyForm.Where(x => x.Id == ProtocolInfo.StudyFormId).Select(x => new { x.Name, x.RodName }).FirstOrDefault();
                    string form = SF.Name + " форма обучения";
                    string form2 = "по " + SF.RodName + " форме";

                    wd.SetFields("Граждан", isRus ? "граждан РФ" : "иностранных граждан");
                    wd.SetFields("Граждан2", isRus ? "граждан Российской Федерации" : "");
                    wd.SetFields("Стипендия", (ProtocolInfo.StudyBasisId == 2 || ProtocolInfo.StudyFormId == 2) ? "" : "\r\nи назначении стипендии");
                    wd.SetFields("Стипендия2", (ProtocolInfo.StudyBasisId == 2 || ProtocolInfo.StudyFormId == 2) ? "" : " и назначении стипендии");
                    wd.SetFields("Факультет", facDat);
                    wd.SetFields("Форма", form);
                    wd.SetFields("Основа", basis);
                    wd.SetFields("БакСпец", StudyLevelName);
                    wd.SetFields("НапрСпец", string.Format(" направлению {0} «{1}»", LicenseProgramCode, LicenseProgramName));
                    wd.SetFields("ПриказОт", docDate);
                    wd.SetFields("ПриказНомер", docNum);
                    wd.SetFields("ПриказОт2", docDate);
                    wd.SetFields("ПриказНомер2", docNum);
                    wd.SetFields("Сокращ", sec);

                    int curRow = 4;
                    var lst = (from extabit in ctx.extAbit
                               join extdisEntryView in ctx.extDisEntryView on extabit.Id equals extdisEntryView.AbiturientId
                               join extperson in ctx.extPerson on extabit.PersonId equals extperson.Id
                               join country in ctx.Country on extperson.NationalityId equals country.Id
                               join competition in ctx.Competition on extabit.CompetitionId equals competition.Id
                               join extabitMarksSum in ctx.extAbitMarksSum on extabit.Id equals extabitMarksSum.Id into extabitMarksSum2
                               from extabitMarksSum in extabitMarksSum2.DefaultIfEmpty()
                               where extdisEntryView.Id == gProtocolId && extdisEntryView.StudyLevelGroupId == MainClass.studyLevelGroupId && (isRus ? extperson.NationalityId == 1 : extperson.NationalityId != 1)
                               orderby extabit.ProfileName, country.NameRod, extperson.FIO
                               select new
                               {
                                   TotalSum = extabitMarksSum.TotalSum,
                                   ФИО = extabit.FIO,
                               }).ToList().Distinct().OrderBy(x => x.ФИО).Select(x =>
                                   new
                                   {
                                       TotalSum = x.TotalSum.ToString(),
                                       ФИО = x.ФИО,
                                   }
                               );

                    foreach (var v in lst)
                    {
                        td.AddRow(1);
                        curRow++;
                        td[0, curRow] = string.Format("\t\tп. № {0} {1} - исключить.", v.ФИО, v.TotalSum);
                    }
                }
            }
            catch (WordException we)
            {
                WinFormsServ.Error(we.Message);
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
        }
        public static void PrintDisEntryView(string protocolId)
        {
            try
            {
                WordDoc wd = new WordDoc(string.Format(@"{0}\DisEntryView.dot", MainClass.dirTemplates));
                TableDoc td = wd.Tables[0];

                Guid gProtocolId = Guid.Parse(protocolId);
                var ProtocolInfo = ProtocolDataProvider.GetProtocolInfo(gProtocolId, 5);

                using (PriemEntities ctx = new PriemEntities())
                {
                    DateTime protocolDate = ProtocolInfo.Date;
                    string protocolNum = ProtocolInfo.Number;

                    Guid entryProtocolId =
                        (from extEntryView in ctx.extEntryView_ForDisEntered
                         join extDisEntryView in ctx.extDisEntryView on extEntryView.AbiturientId equals extDisEntryView.AbiturientId
                         where !extDisEntryView.IsOld && extDisEntryView.Id == gProtocolId
                         select extEntryView.Id).FirstOrDefault();

                    string docNum = "НОМЕР";
                    string docDate = "ДАТА";

                    string facDat =
                        (from protocol in ctx.Protocol
                         join Fac in ctx.SP_Faculty on protocol.FacultyId equals Fac.Id
                         where protocol.Id == gProtocolId
                         select Fac.DatName).FirstOrDefault().ToString();

                    string list = string.Empty, sec = string.Empty;
                    if (ProtocolInfo.IsListener)
                        list = " в качестве слушателя";
                    if (ProtocolInfo.IsSecond)
                        sec = " (для лиц с ВО)";
                    if (ProtocolInfo.IsReduced)
                        sec = " (сокращенной)";

                    string LicenseProgramName =
                        (from entry in ctx.extEntry
                         join extdisEntryView in ctx.extDisEntryView on entry.LicenseProgramId equals extdisEntryView.LicenseProgramId
                         where extdisEntryView.Id == gProtocolId && extdisEntryView.StudyLevelGroupId == MainClass.studyLevelGroupId
                         select entry.LicenseProgramName).FirstOrDefault();

                    string LicenseProgramCode =
                        (from entry in ctx.extEntry
                         join extdisEntryView in ctx.extDisEntryView on entry.LicenseProgramId equals extdisEntryView.LicenseProgramId
                         where extdisEntryView.Id == gProtocolId && extdisEntryView.StudyLevelGroupId == MainClass.studyLevelGroupId
                         select entry.LicenseProgramCode).FirstOrDefault();

                    string StudyLevelName =
                        (from entry in ctx.extEntry
                         join extdisEntryView in ctx.extDisEntryView on entry.LicenseProgramId equals extdisEntryView.LicenseProgramId
                         where extdisEntryView.Id == gProtocolId && extdisEntryView.StudyLevelGroupId == MainClass.studyLevelGroupId
                         select entry.StudyLevelName).FirstOrDefault();

                    string basis = string.Empty;
                    switch (ProtocolInfo.StudyBasisId)
                    {
                        case 1:
                            basis = "обучение за счет средств федерального бюджета";
                            break;
                        case 2:
                            basis = string.Format("по договорам оказания государственной услуги по обучению по основной{0} образовательной программе высшего профессионального образования", sec);
                            break;
                    }

                    var SF = ctx.StudyForm.Where(x => x.Id == ProtocolInfo.StudyFormId).Select(x => new { x.Name, x.RodName }).FirstOrDefault();
                    string form = SF.Name + " форма обучения";
                    string form2 = "по " + SF.RodName + " форме";

                    var lst = (from extabit in ctx.extAbit
                               join extdisEntryView in ctx.extDisEntryView on extabit.Id equals extdisEntryView.AbiturientId
                               join extperson in ctx.extPerson on extabit.PersonId equals extperson.Id
                               join country in ctx.Country on extperson.NationalityId equals country.Id
                               join competition in ctx.Competition on extabit.CompetitionId equals competition.Id
                               join extabitMarksSum in ctx.extAbitMarksSum on extabit.Id equals extabitMarksSum.Id into extabitMarksSum2
                               from extabitMarksSum in extabitMarksSum2.DefaultIfEmpty()
                               where extdisEntryView.Id == gProtocolId && extdisEntryView.StudyLevelGroupId == MainClass.studyLevelGroupId
                               orderby extabit.ProfileName, country.NameRod, extperson.FIO
                               select new
                               {
                                   TotalSum = extabitMarksSum.TotalSum,
                                   ФИО = extabit.FIO,
                                   extperson.NationalityId
                               }).ToList().Distinct().OrderBy(x => x.ФИО).Select(x =>
                                   new
                                   {
                                       TotalSum = x.TotalSum.ToString(),
                                       ФИО = x.ФИО,
                                       x.NationalityId
                                   }
                               );

                    bool isRus = lst.Where(x => x.NationalityId != 1).Count() == 0;

                    wd.SetFields("Граждан", "граждан РФ" + (isRus ? "" : " и иностранных граждан"));
                    wd.SetFields("Граждан2", isRus ? "граждан Российской Федерации" : "");
                    wd.SetFields("Стипендия", ProtocolInfo.StudyBasisId == 2 ? "" : "и назначении стипендии");
                    wd.SetFields("Стипендия2", ProtocolInfo.StudyBasisId == 2 ? "" : "и назначении стипендии");
                    wd.SetFields("Факультет", facDat);
                    wd.SetFields("Форма", form);
                    wd.SetFields("Основа", basis);
                    wd.SetFields("БакСпец", StudyLevelName);
                    wd.SetFields("НапрСпец", string.Format(" направлению {0} «{1}»", LicenseProgramCode, LicenseProgramName));
                    wd.SetFields("ПриказОт", docDate);
                    wd.SetFields("ПриказНомер", docNum);
                    wd.SetFields("ПриказОт2", docDate);
                    wd.SetFields("ПриказНомер2", docNum);
                    wd.SetFields("ПредставлениеОт", protocolDate.ToShortDateString());
                    wd.SetFields("ПредставлениеНомер", protocolNum);
                    wd.SetFields("Сокращ", sec);

                    int curRow = 4;
                    foreach (var v in lst)
                    {
                        td.AddRow(1);
                        curRow++;
                        td[0, curRow] = string.Format("\t\tп. № {0}, {1} - исключить.", v.ФИО, v.TotalSum);
                    }
                }
            }
            catch (WordException we)
            {
                WinFormsServ.Error(we.Message);
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
        }

        public static void PrintDogovor(Guid dogId, Guid abitId, bool forPrint)
        {
            using (PriemEntities context = new PriemEntities())
            {
                var abit = context.extAbit.Where(x => x.Id == abitId).FirstOrDefault();
                if (abit == null)
                {
                    WinFormsServ.Error("Не удалось загрузить данные заявления");
                    return;
                }

                var person = context.extPerson.Where(x => x.Id == abit.PersonId).FirstOrDefault();
                if (person == null)
                {
                    WinFormsServ.Error("Не удалось загрузить данные абитуриента");
                    return;
                }

                var dogovorInfo =
                    (from pd in context.PaidData
                     join pi in context.PayDataEntry on pd.Abiturient.EntryId equals pi.EntryId into pi2
                     from pi in pi2.DefaultIfEmpty()
                     where pd.Id == dogId
                     select new
                     {
                         pd.DogovorNum,
                         DogovorTypeName = pd.DogovorType.Name,
                         pd.DogovorDate,
                         pd.Qualification,
                         pd.Srok,
                         pd.SrokIndividual,
                         pd.DateStart,
                         pd.DateFinish,
                         pd.SumTotal,
                         pd.SumFirstYear,
                         pd.SumFirstPeriod,
                         pd.Parent,
                         Prorector = pd.Prorektor.NameFull,
                         PayPeriodName = pd.PayPeriod.Name,
                         pd.AbitFIORod,
                         pd.AbiturientId,
                         pd.Customer,
                         pd.CustomerLico,
                         pd.CustomerReason,
                         pd.CustomerAddress,
                         pd.CustomerPassport,
                         pd.CustomerPassportAuthor,
                         pd.CustomerINN,
                         pd.CustomerRS,
                         pd.Prorektor.DateDov,
                         pd.Prorektor.NumberDov,
                         PayPeriod = pd.PayPeriod.Name,
                         PayPeriodPad = pd.PayPeriod.NamePad,
                         DogovorTypeId = pd.DogovorTypeId,
                         pi.UniverName,
                         pi.UniverAddress,
                         pi.UniverINN,
                         pi.UniverRS,
                         pi.Props
                     }).FirstOrDefault();

                string dogType = dogovorInfo.DogovorTypeId.ToString();

                WordDoc wd = new WordDoc(string.Format(@"{0}\Dogovor{1}.dot", MainClass.dirTemplates, dogType), !forPrint);

                //вступление
                wd.SetFields("DogovorNum", dogovorInfo.DogovorNum.ToString());
                wd.SetFields("DogovorDate", dogovorInfo.DogovorDate.ToLongDateString());
                
                //проректор и студент
                wd.SetFields("Lico", dogovorInfo.Prorector);
                wd.SetFields("LicoDate", dogovorInfo.DateDov.ToString() + "г.");
                wd.SetFields("LicoNum", dogovorInfo.NumberDov.ToString());
                wd.SetFields("FIO", person.FIO);
                wd.SetFields("Sex", (person.Sex) ? "ый" : "ая");

                string programcode = abit.ObrazProgramCrypt.Trim();
                string profcode = abit.LicenseProgramCode.Trim();

                wd.SetFields("ObrazProgramName", "(" + programcode + ") " + abit.ObrazProgramName.Trim());
                wd.SetFields("Profession", "(" + profcode + ") " + abit.LicenseProgramName);
                wd.SetFields("StudyCourse", "1");
                wd.SetFields("StudyFaculty", abit.FacultyName);
                
                string form = context.StudyForm.Where(x => x.Id == abit.StudyFormId).Select(x => x.Name).FirstOrDefault().ToLower();
                wd.SetFields("StudyForm", form.ToLower());
                
                wd.SetFields("Qualification", dogovorInfo.Qualification);

                //сроки обучения
                wd.SetFields("Srok", dogovorInfo.Srok);

                DateTime dStart = dogovorInfo.DateStart;
                wd.SetFields("DateStart", dStart.ToLongDateString());
                
                DateTime dFinish = dogovorInfo.DateFinish;
                wd.SetFields("DateFinish", dFinish.ToLongDateString());

                //суммы обучения
                wd.SetFields("SumTotal", dogovorInfo.SumTotal);
                wd.SetFields("SumFirstPeriod", dogovorInfo.SumFirstPeriod);//dsRow["SumFirstPeriod"].ToString()
                

                wd.SetFields("Address1", string.Format("{0} {1}, {2}, {3}, ", person.Code, person.CountryName, person.RegionName, person.City));
                wd.SetFields("Address2", string.Format("{0} дом {1} {2} кв. {3}", person.Street, person.House, person.Korpus == string.Empty ? "" : "корп. " + person.Korpus, person.Flat));

                wd.SetFields("Passport", "серия " + person.PassportSeries + " № " + person.PassportNumber);
                wd.SetFields("PassportAuthorDate", person.PassportDate.Value.ToShortDateString());
                wd.SetFields("PassportAuthor", person.PassportAuthor);
                
                wd.SetFields("PhoneNumber", person.Phone + (String.IsNullOrEmpty(person.Mobiles) ? "" : ", доп.: " + person.Mobiles));

                wd.SetFields("UniverName", dogovorInfo.UniverName);
                wd.SetFields("UniverAddress", dogovorInfo.UniverAddress);
                wd.SetFields("UniverINN", dogovorInfo.UniverINN);
                wd.SetFields("Props", dogovorInfo.Props);

                switch (dogType)
                {
                    // обычный
                    case "1":
                        {
                            break;
                        }
                    // физ лицо
                    case "2":
                        {
                            wd.SetFields("CustomerLico", dogovorInfo.Customer);
                            wd.SetFields("CustomerAddress", dogovorInfo.CustomerAddress);
                            wd.SetFields("CustomerINN", "Паспорт: " + dogovorInfo.CustomerPassport);
                            wd.SetFields("CustomerRS", "Выдан: " + dogovorInfo.CustomerPassportAuthor);

                            break;
                        }
                    // мат кап
                    case "4":
                        {
                            wd.SetFields("Customer", dogovorInfo.Customer);
                            wd.SetFields("CustomerAddress", dogovorInfo.CustomerAddress);
                            wd.SetFields("CustomerINN", dogovorInfo.CustomerPassport);
                            wd.SetFields("CustomerRS", dogovorInfo.CustomerPassportAuthor);

                            break;
                        }
                    // юридическое лицо
                    case "3":
                        {
                            wd.SetFields("Customer", dogovorInfo.Customer);
                            wd.SetFields("CustomerLico", dogovorInfo.CustomerLico);
                            wd.SetFields("CustomerReason", dogovorInfo.CustomerReason);
                            wd.SetFields("CustomerAddress", dogovorInfo.CustomerAddress);
                            wd.SetFields("CustomerINN", "ИНН " + dogovorInfo.CustomerINN);
                            wd.SetFields("CustomerRS", "Р/С " + dogovorInfo.CustomerRS);

                            break;
                        }
                }

                if (forPrint)
                {
                    wd.Print();
                    wd.Close();
                }
            }
        }

        public static void PrintDocInventory(IList<int> ids, Guid? _abitId)
        {
            string strIds = Util.BuildStringWithCollection(ids);
            using (PriemEntities context = new PriemEntities())
            {
                var abit = context.extAbit.Where(x => x.Id == _abitId).FirstOrDefault();
                if (abit == null)
                {
                    WinFormsServ.Error("Не найдены данные по заявлению!");
                    return;
                }
                Guid PersonId = abit.PersonId;
                var person = context.Person.Where(x => x.Id == PersonId).FirstOrDefault();
                if (person == null)
                {
                    WinFormsServ.Error("Не найдены данные по человеку!");
                    return;
                }
                string FIO = (person.Surname ?? "") + " " + (person.Name ?? "") + " " + (person.SecondName ?? "");
                WordDoc wd = new WordDoc(string.Format(@"{0}\DocInventory.dot", MainClass.dirTemplates), true);

                wd.SetFields("FIO", FIO);

                var docs = context.AbitDoc.Join(ids, x => x.Id, y => y, (x, y) => new { x.Id, x.Name }).Select(x => x.Name);

                int i = 1;
                wd.AddNewTable(docs.Count(), 1);
                foreach (var d in docs)
                {
                    wd.Tables[0][0, i - 1] = i.ToString() + ") " + d + "\n";
                    i++;
                }
            }
        }

        public static void PrintRatingProtocol(int? iStudyFormId, int? iStudyBasisId, int? iFacultyId, int? iLicenseProgramId, int? iObrazProgramId, int? iProfileId, bool isCel, bool isCrimea, int plan, string savePath, bool isSecond, bool isReduced, bool isParallel, bool isQuota)
        {
            FileStream fileS = null;
            try
            {
                Guid fixId;
                int? docNum;
                string form;
                string facDat;
                string prof;
                string obProg;
                string spec;

                using (PriemEntities ctx = new PriemEntities())
                {
                    fixId = (from fixierenView in ctx.FixierenView
                             where fixierenView.StudyFormId == iStudyFormId && fixierenView.StudyBasisId == iStudyBasisId && fixierenView.FacultyId == iFacultyId && fixierenView.LicenseProgramId == iLicenseProgramId &&
                             fixierenView.ObrazProgramId == iObrazProgramId && (iProfileId.HasValue ? fixierenView.ProfileId == iProfileId : true) && fixierenView.IsCel == isCel && fixierenView.IsCrimea == isCrimea && fixierenView.IsSecond == isSecond && fixierenView.IsParallel == isParallel && fixierenView.IsReduced == isReduced && fixierenView.IsQuota == isQuota
                             select fixierenView.Id).FirstOrDefault();

                    docNum = (from fixierenView in ctx.FixierenView
                              where fixierenView.Id == fixId
                              select fixierenView.DocNum).FirstOrDefault();

                    form = (from studyForm in ctx.StudyForm
                            where studyForm.Id == iStudyFormId
                            select studyForm.Acronym).FirstOrDefault();

                    facDat = (from sP_Faculty in ctx.SP_Faculty
                              where sP_Faculty.Id == iFacultyId
                              select sP_Faculty.DatName).FirstOrDefault();

                    prof = (from entry in ctx.Entry
                            where entry.LicenseProgramId == iLicenseProgramId
                            select entry.SP_LicenseProgram.Code + " " + entry.SP_LicenseProgram.Name).FirstOrDefault();

                    obProg = (from entry in ctx.Entry
                              where entry.ObrazProgramId == iObrazProgramId
                              select entry.StudyLevel.Acronym + "." + entry.SP_ObrazProgram.Number + "." + MainClass.sPriemYear + " " + entry.SP_ObrazProgram.Name).FirstOrDefault();

                    spec = (from entry in ctx.Entry
                            where iProfileId.HasValue ? entry.ProfileId == iProfileId : entry.ProfileId == null
                            select entry.SP_Profile.Name).FirstOrDefault();
                }

                string basis = string.Empty;

                switch (iStudyBasisId)
                {
                    case 1:
                        basis = "обучение за счет средств федерального бюджета";
                        break;
                    case 2:
                        basis = "обучение по договорам с оплатой стоимости обучения";
                        break;
                }

                Document document = new Document(PageSize.A4.Rotate(), 50, 50, 50, 50);

                using (fileS = new FileStream(savePath, FileMode.Create))
                {

                    BaseFont bfTimes = BaseFont.CreateFont(string.Format(@"{0}\times.ttf", MainClass.dirTemplates), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    Font font = new Font(bfTimes, 12);

                    PdfWriter writer = PdfWriter.GetInstance(document, fileS);
                    document.Open();

                    float firstLineIndent = 30f;
                    //HEADER
                    Paragraph p = new Paragraph("ПРАВИТЕЛЬСТВО РОССИЙСКОЙ ФЕДЕРАЦИИ", new Font(bfTimes, 12, Font.BOLD));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    p = new Paragraph("ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ВЫСШЕГО", new Font(bfTimes, 10));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    p = new Paragraph("ПРОФЕССИОНАЛЬНОГО ОБРАЗОВАНИЯ", new Font(bfTimes, 10));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    p = new Paragraph("САНКТ-ПЕТЕРБУРГСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ", new Font(bfTimes, 12, Font.BOLD));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    p = new Paragraph("(СПбГУ)", new Font(bfTimes, 12, Font.BOLD));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    p = new Paragraph("ПРЕДСТАВЛЕНИЕ", new Font(bfTimes, 20, Font.BOLD));
                    p.Alignment = Element.ALIGN_CENTER;
                    document.Add(p);

                    p = new Paragraph(10f);
                    p.Add(new Paragraph("По " + facDat, font));
                    p.Add(new Paragraph((form + " форма обучения").ToLower(), font));
                    p.Add(new Paragraph(basis, font));
                    p.IndentationLeft = 510;
                    document.Add(p);

                    p = new Paragraph("О зачислении на 1 курс", font);
                    p.SpacingBefore = 10f;
                    document.Add(p);

                    p = new Paragraph(@"В соответствии с Федеральным законом от 22.08.1996 N 125-Ф3 (ред. от 21.12.2009) «О высшем и послевузовском профессиональном образовании», Порядком приема граждан в имеющие государственную аккредитацию образовательные учреждения высшего профессионального образования, утвержденным Приказом Министерства образования и науки Российской Федерации от 21.10.2009 N 442 (ред. от 11.05.2010)", font);
                    p.SpacingBefore = 10f;
                    p.Alignment = Element.ALIGN_JUSTIFIED;
                    p.FirstLineIndent = firstLineIndent;
                    document.Add(p);

                    p = new Paragraph("Представляем на рассмотрение Приемной комиссии СПбГУ полный пофамильный перечень поступающих на 1 курс обучения по основным образовательным программам высшего профессионального образования:", font);
                    p.FirstLineIndent = firstLineIndent;
                    p.Alignment = Element.ALIGN_JUSTIFIED;
                    p.SpacingBefore = 20f;
                    document.Add(p);

                    p = new Paragraph("по направлению " + prof, font);
                    p.FirstLineIndent = firstLineIndent * 2;
                    document.Add(p);

                    p = new Paragraph("по образовательной программе " + obProg, font);
                    p.FirstLineIndent = firstLineIndent * 2;
                    document.Add(p);

                    if (!string.IsNullOrEmpty(spec))
                    {
                        p = new Paragraph("по профилю " + spec, font);
                        p.FirstLineIndent = firstLineIndent * 2;
                        document.Add(p);
                    }

                    //Table
                    float[] headerwidths = { 5, 9, 9, 19, 6, 10, 10, 7, 11, 14 };

                    PdfPTable t = new PdfPTable(10);
                    t.SetWidthPercentage(headerwidths, document.PageSize);
                    t.WidthPercentage = 100f;
                    t.SpacingBefore = 10f;
                    t.SpacingAfter = 10f;

                    t.HeaderRows = 1;

                    string[] headers = new string[]
                    {
                        "№ п/п",
                        "Рег. номер",
                        "Ид. номер",
                        "ФИО",
                        "Cумма баллов",
                        "Подлинники документов",
                        "Рейтинговый коэффициент",
                        "Конкурс",
                        "Профильное вступительное испытание",
                        "Дополнительное вступительное испытание"
                    };
                    foreach (string h in headers)
                    {
                        PdfPCell cell = new PdfPCell();
                        cell.BorderColor = Color.BLACK;
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                        cell.AddElement(new Phrase(h, new Font(bfTimes, 12, Font.BOLD)));

                        t.AddCell(cell);
                    }

                    int counter = 0;

                    using (PriemEntities ctx = new PriemEntities())
                    {
                        var lst = (from extabit in ctx.extAbit
                                   join fixieren in ctx.Fixieren on extabit.Id equals fixieren.AbiturientId
                                   join fixierenView in ctx.FixierenView on fixieren.FixierenViewId equals fixierenView.Id into fixierenView2
                                   from fixierenView in fixierenView2.DefaultIfEmpty()
                                   join extperson in ctx.extPerson on extabit.PersonId equals extperson.Id
                                   join competition in ctx.Competition on extabit.CompetitionId equals competition.Id
                                   join hlpabiturientProfAdd in ctx.hlpAbiturientProfAdd on extabit.Id equals hlpabiturientProfAdd.Id into hlpabiturientProfAdd2
                                   from hlpabiturientProfAdd in hlpabiturientProfAdd2.DefaultIfEmpty()
                                   join hlpabiturientProf in ctx.hlpAbiturientProf on extabit.Id equals hlpabiturientProf.Id into hlpabiturientProf2
                                   from hlpabiturientProf in hlpabiturientProf2.DefaultIfEmpty()
                                   join extabitMarksSum in ctx.extAbitMarksSum on extabit.Id equals extabitMarksSum.Id into extabitMarksSum2
                                   from extabitMarksSum in extabitMarksSum2.DefaultIfEmpty()
                                   where fixierenView.Id == fixId
                                   orderby fixieren.Number
                                   select new
                                   {
                                       Id = extabit.Id,
                                       Рег_Номер = extabit.RegNum,
                                       Ид_номер = extabit.PersonNum,
                                       ФИО = extabit.FIO,
                                       Сумма_баллов = extabitMarksSum.TotalSum,
                                       Кол_во_оценок = extabitMarksSum.TotalCount,
                                       Подлинники_документов = extabit.HasOriginals ? "Да" : "Нет",
                                       Рейтинговый_коэффициент = extabit.Coefficient,
                                       Конкурс = competition.Name,
                                       Проф_экзамен = hlpabiturientProf.Prof,
                                       Доп_экзамен = hlpabiturientProfAdd.ProfAdd,
                                       comp = competition.Id == 1 ? 1 : (competition.Id == 2 || competition.Id == 7) && extperson.Privileges > 0 ? 2 : 3,
                                       noexamssort = competition.Id == 1 ? extabit.Coefficient : 0
                                   }).ToList().Distinct().Select(x =>
                                       new
                                       {
                                           Id = x.Id.ToString(),
                                           Рег_Номер = x.Рег_Номер,
                                           Ид_номер = x.Ид_номер,
                                           ФИО = x.ФИО,
                                           Сумма_баллов = x.Сумма_баллов,
                                           Кол_во_оценок = x.Кол_во_оценок,
                                           Подлинники_документов = x.Подлинники_документов,
                                           Рейтинговый_коэффициент = x.Рейтинговый_коэффициент,
                                           Конкурс = x.Конкурс,
                                           Проф_экзамен = x.Проф_экзамен,
                                           Доп_экзамен = x.Доп_экзамен,
                                           comp = x.comp,
                                           noexamssort = x.noexamssort
                                       }
                                   );

                        foreach (var v in lst)
                        {
                            ++counter;
                            t.AddCell(new Phrase(counter.ToString(), font));
                            t.AddCell(new Phrase(v.Рег_Номер, font));
                            t.AddCell(new Phrase(v.Ид_номер, font));
                            t.AddCell(new Phrase(v.ФИО, font));
                            t.AddCell(new Phrase(v.Сумма_баллов.ToString(), font));
                            t.AddCell(new Phrase(v.Подлинники_документов, font));
                            t.AddCell(new Phrase(v.Рейтинговый_коэффициент.ToString(), font));
                            t.AddCell(new Phrase(v.Конкурс, font));
                            t.AddCell(new Phrase(v.Проф_экзамен.ToString(), font));
                            t.AddCell(new Phrase(v.Доп_экзамен.ToString(), font));
                        }
                    }

                    document.Add(t);

                    //FOOTER
                    p = new Paragraph();
                    p.SpacingBefore = 30f;
                    p.Alignment = Element.ALIGN_JUSTIFIED;
                    p.FirstLineIndent = firstLineIndent;
                    p.Add(new Phrase("Основание:", new Font(bfTimes, 12, Font.BOLD)));
                    p.Add(new Phrase(" личные заявления, результаты вступительных испытаний, документы, подтверждающие право на поступление без вступительных испытаний или внеконкурсное зачисление.", font));
                    document.Add(p);


                    p = new Paragraph(30f);
                    p.KeepTogether = true;
                    p.Add(new Paragraph("Ответственный секретарь по приему документов по группе направлений:", font));
                    p.Add(new Paragraph("Заместитель начальника управления - советник проректора по группе направлений:", font));
                    //p.Add(new Paragraph("Ответственный секретарь приемной комиссии:", font));

                    document.Add(p);

                    p = new Paragraph(30f);
                    p.Add(new Phrase("В." + iFacultyId.ToString() + "." + docNum, font));
                    document.Add(p);
                    document.Close();

                    Process pr = new Process();

                    pr.StartInfo.Verb = "Open";
                    pr.StartInfo.FileName = string.Format(savePath);
                    pr.Start();

                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message);
            }
            finally
            {
                if (fileS != null)
                    fileS.Dispose();
            }
        }

        public static string[] GetSplittedStrings(string sourceStr, int firstStrLen, int strLen, int numOfStrings)
        {
            sourceStr = sourceStr ?? "";
            string[] retStr = new string[numOfStrings];
            int index = 0, startindex = 0;
            for (int i = 0; i < numOfStrings; i++)
            {
                if (sourceStr.Length > startindex && startindex >= 0)
                {
                    int rowLength = firstStrLen;//длина первой строки
                    if (i > 1) //длина остальных строк одинакова
                        rowLength = strLen;
                    index = startindex + rowLength;
                    if (index < sourceStr.Length)
                    {
                        index = sourceStr.IndexOf(" ", index);
                        string val = index > 0 ? sourceStr.Substring(startindex, index - startindex) : sourceStr.Substring(startindex);
                        retStr[i] = val;
                    }
                    else
                        retStr[i] = sourceStr.Substring(startindex);
                }
                startindex = index;
            }

            return retStr;
        }
        
        public static byte[] MergePdfFiles(List<byte[]> lstFilesBinary)
        {
            MemoryStream ms = new MemoryStream();
            Document document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.GetInstance(document, ms);

            document.Open();

            foreach (byte[] doc in lstFilesBinary)
            {
                PdfReader reader = new PdfReader(doc);
                int n = reader.NumberOfPages;
                //writer.SetEncryption(PdfWriter.STRENGTH128BITS, "", "", PdfWriter.ALLOW_SCREENREADERS | PdfWriter.ALLOW_PRINTING | PdfWriter.AllowPrinting);

                PdfContentByte cb = writer.DirectContent;
                PdfImportedPage page;

                for (int i = 0; i < n; i++)
                {
                    document.NewPage();
                    page = writer.GetImportedPage(reader, i + 1);
                    cb.AddTemplate(page, 1f, 0, 0, 1f, 0, 0);
                }
            }

            document.Close();
            return ms.ToArray();
        }
    }

    public class ShortAppcationDetails
    {
        public Guid ApplicationId { get; set; }
        public int? CurrVersion { get; set; }
        public DateTime? CurrDate { get; set; }

        public string ObrazProgramName { get; set; }
        public string ProfileName { get; set; }
        public int Priority { get; set; }
    }
    public class ShortAppcation
    {
        public Guid ApplicationId { get; set; }
        public int Priority { get; set; }
        public string LicenseProgramName { get; set; }
        public string ObrazProgramName { get; set; }
        public string ProfileName { get; set; }

        public bool HasInnerPriorities { get; set; }
        public int InnerPrioritiesNum { get; set; }

        public int StudyFormId { get; set; }
        public int StudyBasisId { get; set; }
    }
}
