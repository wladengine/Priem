using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.Objects;
using System.Data;
using System.Transactions;

using BDClassLib;
using EducServLib;

namespace Priem
{
    public static class SomeMethodsClass
    {
        public static void FillOlymps()
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.Filter = "CSV|*.csv";
            ofd.Multiselect = true;

            if (!(ofd.ShowDialog() == DialogResult.OK))
                return;
            foreach (string fileName in ofd.FileNames)
            {
                using (StreamReader sr = new StreamReader(fileName, Encoding.GetEncoding(1251)))
                {
                    List<string[]> list = new List<string[]>();

                    while (!sr.EndOfStream)
                    {
                        string str = sr.ReadLine();
                        char[] split = new char[] { ';' };
                        string[] values = str.Split(split, 4);

                        list.Add(values);
                    }

                    int i = 0;
                    foreach (string[] val in list)
                    {
                        i++;

                        if (val.Length < 4)
                        {
                            MessageBox.Show("Строка " + i.ToString() + ": некорректный формат строки!");
                            continue;
                        }

                        string num = val[0].Trim();
                        int j;
                        int? numInt;
                        if(int.TryParse(num, out j))
                            numInt = j;
                        else
                            numInt = null;
                       
                        string name = val[1].Trim();
                        string olName = num + " - " + name;                       

                        using (PriemEntities context = new PriemEntities())
                        {
                            int? olNameId;
                            var ol = from qq in context.OlympName
                                       where qq.Name == olName
                                       select qq;

                            if (ol != null && ol.Count() > 0)
                                olNameId = (ol.First()).Id;
                            else
                            {
                                ObjectParameter olEnt = new ObjectParameter("id", typeof(Int32));
                                context.OlympName_Insert(olName, numInt, olEnt);
                                olNameId = (int)olEnt.Value;
                            }

                            string subjName = val[2].Trim();
                            
                            int? subjId;
                            var subj = from qq in context.OlympSubject
                                       where qq.Name == subjName
                                       select qq;

                            if (subj != null && subj.Count() > 0)
                                subjId = (subj.First()).Id;
                            else
                            {
                                ObjectParameter sbEnt = new ObjectParameter("id", typeof(Int32));
                                context.OlympSubject_Insert(subjName, "", sbEnt);
                                subjId = (int)sbEnt.Value;
                            } 

                            int? levId;
                            int level;
                            if (!int.TryParse(val[3].Trim(), out level))
                                levId = null;
                            else
                                levId = level;

                            int cnt = (from ob in context.OlympBook
                                       where ob.OlympTypeId == 4 && ob.OlympNameId == olNameId && ob.OlympSubjectId == subjId && ob.OlympLevelId == levId
                                       select ob).Count();

                            if (cnt > 0)
                                continue;
                            
                            ObjectParameter EntId = new ObjectParameter("id", typeof(Int32));
                            context.OlympBook_Insert(4, olNameId, subjId, levId, EntId);                        
                        }
                    }
                }
            }

            MessageBox.Show("Выполнено!");
        }

        // проверка на уникальность абитуриента
        public static bool CheckThreeAbits(PriemEntities context, Guid? personId, int? LicenseProgramId)
        {
            // если прием - то проверяем на три заявления
            if (MainClass.dbType != PriemType.Priem)
                return true;  

            //просто сосчитаем количество созданных активных конкурсов на человека
            var concurses = (from allab in context.Abiturient
                             where allab.PersonId == personId
                                 && allab.Entry.LicenseProgramId != LicenseProgramId
                                 //&& allab.ObrazProgramId != ObrazProgramId
                                 //&& (ProfileId == null ? allab.ProfileId != null : allab.ProfileId != ProfileId)
                             && allab.Entry.StudyLevel.LevelGroupId == MainClass.studyLevelGroupId
                             && allab.BackDoc != true
                             && !allab.IsGosLine //считать гослинию нам не надо, только не-гослинию
                             select new { allab.Entry.LicenseProgramId }).Distinct();
            return (concurses.Count() < 3);
        }

        // проверка на уникальность абитуриента
        private static bool CheckIdent(extPerson person, string AttestatSeries, string AttestatNum)
        {
            using (PriemEntities context = new PriemEntities())
            {
                ObjectParameter boolPar = new ObjectParameter("result", typeof(bool));

                context.CheckPersonIdent(person.Surname, person.Name, person.SecondName, person.BirthDate, person.PassportSeries, person.PassportNumber,
                    AttestatSeries, AttestatNum, boolPar);
               
                return Convert.ToBoolean(boolPar.Value);
            }
        }

        public static void SetBackDocForBudgetInEntryView()
        {
            if (!MainClass.IsPasha())
                return;

            using (PriemEntities context = new PriemEntities())
            {
                if (MessageBox.Show("Проставить 'Забрал документы' для платных заявлений, поступивших на бесплатное?", "Внимание", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {                    
                    using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew, TimeSpan.FromHours(1)))
                    {
                        var abits = from ev in context.extEntryView
                                    join ab in context.extAbit
                                    on ev.AbiturientId equals ab.Id
                                    where ab.StudyLevelGroupId == 1 && ab.StudyBasisId == 1 && !ab.BackDoc && ab.HasOriginals
                                    select ab;

                        foreach (extAbit abit in abits)
                        {
                            var abBackDocks = from ab in context.extAbit
                                              where ab.StudyLevelGroupId == abit.StudyLevelGroupId
                                              && ab.IsReduced == abit.IsReduced && ab.IsParallel == abit.IsParallel && ab.IsSecond == abit.IsSecond
                                              && ab.FacultyId == abit.FacultyId && ab.LicenseProgramId == abit.LicenseProgramId
                                              && ab.ObrazProgramId == abit.ObrazProgramId
                                              && (abit.ProfileId == null ? ab.ProfileId == null : ab.ProfileId == abit.ProfileId)
                                              && ab.StudyFormId == abit.StudyFormId
                                              && ab.StudyBasisId == 2
                                              select ab;

                            if (abBackDocks.Count() > 0)
                            {
                                foreach(extAbit abBack in abBackDocks)
                                {
                                    context.Abiturient_UpdateBackDoc(true, DateTime.Now, abBack.Id);
                                }
                            }
                        }


                        transaction.Complete();
                    }
                }
            }
        }

        public static void DeleteDogFromFirstWave()
        {
            if (!MainClass.IsPasha())
                return;

            using (PriemEntities context = new PriemEntities())
            {
                if (MessageBox.Show("Удалить платников, забравших документы или зачисленных на 1 курс, из FirstWave?", "Внимание", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {                    
                    using (TransactionScope transaction = new TransactionScope(TransactionScopeOption.RequiresNew, TimeSpan.FromHours(1)))
                    {
                        context.FirstWave_DELETE_DogEntryBack();

                        transaction.Complete();
                    }
                }
            }
        }

        public static void SplitEntryViews()
        {
            try
            {
                using (PriemEntities context = new PriemEntities())
                {
                    var entryViewNums =
                        (from exEV in context.extEntryView
                         join Ab in context.Abiturient on exEV.AbiturientId equals Ab.Id
                         where Ab.CompetitionId == 1
                         select exEV.Number).Intersect
                        (from exEV in context.extEntryView
                         join Ab in context.Abiturient on exEV.AbiturientId equals Ab.Id
                         where Ab.CompetitionId == 2
                         select exEV.Number);

                    foreach (string EVNumber in entryViewNums)
                    {

                        var abitIds =
                            (from extEV in context.extEntryView
                             join Ab in context.Abiturient on extEV.AbiturientId equals Ab.Id
                             where Ab.CompetitionId == 2 && extEV.Number == EVNumber
                             select Ab.Id).ToList();

                        var protocolInfo = context.extEntryView.Where(x => x.Number == EVNumber)
                            .Select(x => new
                        {
                            x.StudyLevelGroupId,
                            x.FacultyId,
                            x.LicenseProgramId,
                            x.StudyFormId,
                            x.StudyBasisId,
                            x.Date,
                            x.Reason,
                            x.IsSecond,
                            x.IsReduced,
                            x.IsParallel,
                            x.IsListener,
                            x.EntryHeaderId,
                            x.Id
                        }).FirstOrDefault();

                        ObjectParameter idParam = new ObjectParameter("id", typeof(Guid));
                        context.Protocol_Insert(protocolInfo.StudyLevelGroupId, protocolInfo.FacultyId, protocolInfo.LicenseProgramId, protocolInfo.StudyFormId, protocolInfo.StudyBasisId, EVNumber + "\\1",
                            protocolInfo.Date, 4, protocolInfo.Reason, false, null, protocolInfo.IsSecond, protocolInfo.IsListener, idParam);

                        Guid protocolId = (Guid)idParam.Value;

                        foreach (var id in abitIds)
                        {
                            context.ProtocolHistory_Insert(id, protocolId, false, null, protocolInfo.EntryHeaderId, idParam);
                            context.ProtocolHistory_Delete(id, protocolInfo.Id);
                        }
                    }
                }
            }
            catch (Exception exc)
            {
                WinFormsServ.Error(exc.Message + (exc.InnerException == null ? "" : "\nINNER EXCEPTION: " + exc.InnerException.Message));
            }
        }

        public static void SetBackDockForCommit(int applicationCommitNumber, List<Guid> AppList)
        {
            using (PriemEntities context = new PriemEntities())
            {
                string message = "Проставить отказ от участия в конкурсе по следующим позициям:";
                int incrmntr = 1;
                var Orig_AppList = new Guid[AppList.Count];
                AppList.CopyTo(Orig_AppList);
                foreach (var App in Orig_AppList)
                {
                    var data = context.Abiturient.Where(x => x.Id == App && !x.BackDoc).Select(x => new
                    {
                        LP = x.Entry.SP_LicenseProgram.Code + " " + x.Entry.SP_LicenseProgram.Name,
                        OP = x.Entry.StudyLevel.Acronym + "." + x.Entry.SP_ObrazProgram.Number + "." + MainClass.sPriemYear + " " + x.Entry.SP_ObrazProgram.Name,
                        Prof = x.Entry.SP_Profile.Name,
                        StudyForm = x.Entry.StudyForm.Acronym,
                        StudyBasis = x.Entry.StudyBasis.Acronym
                    }).FirstOrDefault();

                    if (data == null)//бред, конечно, но кто знает
                        AppList.Remove(App);
                    else
                        message += "\n" + incrmntr++ + ") " + data.LP + "\n" + data.OP + "; Профиль:" + data.Prof + "; " + data.StudyForm + "; " + data.StudyBasis;
                }
                if (AppList.Count == 0)
                {
                    var abit = (from ab in context.Abiturient
                                where ab.CommitNumber == applicationCommitNumber
                                select ab).FirstOrDefault();
                    if (abit != null)
                    {
                        string persId = abit.PersonId.ToString();

                        var dialogres = MessageBox.Show("Открыть карточку абитуриента?", "Внимание", MessageBoxButtons.YesNo);
                        if (dialogres == System.Windows.Forms.DialogResult.Yes)
                            MainClass.OpenCardPerson(persId, null, null);
                    }
                }
                else
                {
                    var dr = MessageBox.Show(message, "Отзыв старого заявления", MessageBoxButtons.YesNo);
                    if (dr == DialogResult.Yes)
                    {
                        dr = MessageBox.Show("У меня есть на руках заявление об отказе от участия в конкурсе", "Внимание!", MessageBoxButtons.YesNo);
                        if (dr == DialogResult.Yes)
                        {
                            foreach (var App in AppList)
                            {
                                context.Abiturient_UpdateBackDoc(true, DateTime.Now, App);
                            }
                        }
                        else //тогда открываем карточку абитуриента
                        {
                            var abit = (from ab in context.Abiturient
                                        where ab.CommitNumber == applicationCommitNumber
                                        select ab).FirstOrDefault();
                            if (abit != null)
                            {
                                string persId = abit.PersonId.ToString();

                                dr = MessageBox.Show("Открыть карточку абитуриента?", "Внимание", MessageBoxButtons.YesNo);
                                if (dr == System.Windows.Forms.DialogResult.Yes)
                                    MainClass.OpenCardPerson(persId, null, null);
                            }
                        }
                    }
                    else //тогда открываем карточку абитуриента
                    {
                        extAbit abit = (from ab in context.extAbit
                                        where ab.CommitNumber == applicationCommitNumber
                                        select ab).FirstOrDefault();

                        if (abit != null)
                        {
                            string fio = abit.FIO;
                            string num = abit.PersonNum;
                            string persId = abit.PersonId.ToString();

                            dr = MessageBox.Show(string.Format("Абитуриент {0} с данным штрих-кодом уже импортирован в базу.\nОткрыть карточку абитуриента?", fio), "Внимание", MessageBoxButtons.YesNo);
                            if (dr == System.Windows.Forms.DialogResult.Yes)
                                MainClass.OpenCardPerson(persId, null, null);
                        }
                    }
                }
            }
        }
    }
}
