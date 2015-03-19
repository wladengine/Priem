using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Linq;

using EducServLib;
using BDClassLib;
using BaseFormsLib;

namespace Priem
{
    public partial class PersonInetList : BookList
    {
        private DBPriem bdcInet;
        private LoadFromInet loadClass;

        //конструктор
        public PersonInetList()
        {
            InitializeComponent();

            Dgv = dgvAbiturients;
            _tableName = "ed.extPerson";
            _title = "Список абитуриентов СПбГУ";

            InitControls();
        }        
        
        //дополнительная инициализация контролов
        protected override void  ExtraInit()
        {
            base.ExtraInit();            

            if (MainClass.RightsJustView())
            {
                btnLoad.Enabled = false;
                btnAdd.Enabled = false;
            }
           
            if (MainClass.dbType == PriemType.PriemMag)
                tbPersonNum.Visible = lblBarcode.Visible = btnLoad.Visible = false;

            //Dgv.CellDoubleClick -= new System.Windows.Forms.DataGridViewCellEventHandler(Dgv_CellDoubleClick);
        }

        //поле поиска
        private void tbSearch_TextChanged(object sender, EventArgs e)
        {
            WinFormsServ.Search(this.dgvAbiturients, "FIO", tbSearch.Text);
        }

        protected override void OpenCard(string id, BaseFormEx formOwner, int? index)
        {
            MainClass.OpenCardPerson(id, formOwner, index);
        }

        protected override void GetSource()
        {
            _sQuery = @"SELECT DISTINCT extPerson.Id, extPerson.PersonNum, extPerson.FIO, extPerson.PassportData, EducInfo.EducDocument FROM ed.extPerson 
LEFT JOIN ed.extPerson_EducationInfo_Current AS EducInfo ON EducInfo.PersonId = extPerson.Id";
            string join = "";

            if (!chbShowAll.Checked)
            {
                join = string.Format(@" LEFT JOIN ed.Abiturient ON Abiturient.PersonId = extPerson.Id 
LEFT JOIN ed.Entry ON Entry.Id=Abiturient.EntryId 
LEFT JOIN ed.StudyLevel ON StudyLevel.Id=Entry.StudyLevelId 
WHERE Abiturient.IsGosLine <> 1 {1} 
AND (StudyLevel.LevelGroupId <> {0} OR StudyLevel.LevelGroupId IS NULL OR StudyLevel.LevelGroupId = {2}) ",
                    MainClass.dbType == PriemType.Priem ? 2 : 1,
                    MainClass.dbType == PriemType.Priem ? "" : "AND EducInfo.SchoolTypeId=4",
                    MainClass.dbType == PriemType.Priem ? 1 : 2);
            }
            
            HelpClass.FillDataGrid(Dgv, _bdc, _sQuery + join, "", " ORDER BY FIO");
            SetVisibleColumnsAndNameColumns();    
        }

        protected override void SetVisibleColumnsAndNameColumns()
        {
            Dgv.AutoGenerateColumns = false;

            foreach (DataGridViewColumn col in Dgv.Columns)
            {
                col.Visible = false;
            }
            
            this.Width = 608;
            dgvAbiturients.Columns["PersonNum"].Width = 70;
            dgvAbiturients.Columns["FIO"].Width = 246;

            SetVisibleColumnsAndNameColumnsOrdered("PersonNum", "Ид_номер", 0);
            SetVisibleColumnsAndNameColumnsOrdered("FIO", "ФИО", 1);
            SetVisibleColumnsAndNameColumnsOrdered("PassportData", "Паспортные данные", 2);
            SetVisibleColumnsAndNameColumnsOrdered("EducDocument", "Документ об образовании", 3);          
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            loadClass = new LoadFromInet();
            bdcInet = loadClass.BDCInet;

            if (!bdcInet.IsOpen)
                return;

            int fileNum = 0;

            string barcText = tbPersonNum.Text.Trim();

            if (barcText == string.Empty)
            {
                WinFormsServ.Error("Не введен номер");
                return;
            }

            if (barcText.Length == 6)
            {
                if (barcText.StartsWith("2") && MainClass.dbType == PriemType.Priem)
                {
                    WinFormsServ.Error("Выбран человек, подавший заявления в магистратуру");
                    return;
                }

                barcText = barcText.Substring(1);
            }

            if (!int.TryParse(barcText, out fileNum))
            {
                WinFormsServ.Error("Неправильно введен номер");
                return;
            }

            //сперва проверим, есть ли ещё "живое" заявление, или абитуриент уже успел его удалить 
            string query = "SELECT COUNT(Id) FROM Abiturient WHERE ApplicationCommitNumber=@Barcode";
            bool bHasInInetBase = (int)bdcInet.GetValue(query, new SortedList<string, object>() { { "@Barcode", fileNum } }) > 0;

            //если заявления нет, то надо проверить, нет ли его в нашей базе
            if (!bHasInInetBase)
            {
                if (MainClass.CheckExistenseAbitCommitNumberInWorkBase(fileNum))
                {
                    //если по данному номеру что-то в рабочей базе есть, то проводим процедуру забора документов
                    using (PriemEntities context = new PriemEntities())
                    {
                        var abits = (from ab in context.Abiturient
                                     where ab.CommitNumber == fileNum
                                     select ab.Id).ToList();

                        SomeMethodsClass.SetBackDockForCommit(fileNum, abits);
                    }
                }
                else
                {
                    //проверить в логах, не сохранилось ли чего-то
                    //если в логах есть, то вытащить оттуда PersonId
                    query = "SELECT PersonId FROM Application_LOG INNER JOIN ApplicationCommit ON Application_LOG.CommitId=ApplicationCommit.Id WHERE ApplicationCommit.IntNumber=@Barcode";
                    Guid? PersonId = (Guid?)bdcInet.GetValue(query, new SortedList<string, object>() { { "@Barcode", fileNum } });
                    if (PersonId.HasValue)
                    {
                        //найти в заявлениях, нет ли нового коммита
                        query = "SELECT COUNT(*) FROM Abiturient WHERE ApplicationCommitNumber<>@Barcode AND PersonId=@PersonId AND StudyLevelGroupId=@SLGrId";
                        bool bHasAnotherInInetBase = (int)bdcInet.GetValue(query,
                            new SortedList<string, object>() { { "@Barcode", fileNum }, { "@PersonId", PersonId.Value }, { "@SLGrId", MainClass.studyLevelGroupId } }) > 0;

                        if (bHasAnotherInInetBase)
                        {
                            //если коммит есть
                            MessageBox.Show("Данное заявление было отозвано в Личном Кабинете. Взамен него было создано новое. Свяжитесь с абитуриентом для выявления последней версии заявления.", "Внимание, данное заявление было отозвано", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        else
                        {
                            //если коммита нет, то ошибка, ищите возможности для восстановления
                            MessageBox.Show("Данное заявление было отозвано в Личном Кабинете. Взамен него другого заявления не создано. Свяжитесь с абитуриентом.", "Внимание, данное заявление было отозвано", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    //если в логах ничего нет, то беда
                    else
                    {
                        MessageBox.Show("Данное заявление было отозвано в Личном Кабинете. Взамен него другого заявления не создано. Свяжитесь с абитуриентом. Восстановить заявление невозможно.", "Внимание, данное заявление было отозвано", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                //проверяем, есть ли в базе заявления по данному номеру коммита
                if (!MainClass.CheckExistenseAbitCommitNumberInWorkBase(fileNum))
                {
                    //если нет, то нужно проверить, есть ли заявления по другому коммиту
                    try
                    {
                        using (PriemEntities context = new PriemEntities())
                        {
                            //для этого надо сперва узнать баркод человека
                            query = "SELECT Person.Barcode FROM Abiturient INNER JOIN Person ON Person.Id = Abiturient.PersonId WHERE ApplicationCommitNumber=@Barcode";
                            int? iPersBarc = (int?)bdcInet.GetValue(query, new SortedList<string, object>() { { "@Barcode", fileNum } });

                            var abits = (from ab in context.Abiturient
                                         where ab.CommitNumber != fileNum
                                         && !ab.BackDoc
                                         && ab.Entry.StudyLevel.LevelGroupId == MainClass.studyLevelGroupId
                                         && iPersBarc.HasValue ? ab.Person.Barcode == iPersBarc : false//если баркод не найдён в оригинальной базе, то это совсем беда
                                         select ab.Id).ToList();

                            //если заявления есть, то нужно осуществить их последовательное проставление "забрал документы"
                            if (abits.Count() > 0)
                            {
                                var dr = MessageBox.Show("У абитуриента имеются активные заявления более ранних версий. Проставить по ним отказ от участия в конкурсе?\nДа - проставить, Нет - не проставить, загрузить новое заявление без отказа по старым", "Внимание", MessageBoxButtons.YesNo);
                                if (dr == System.Windows.Forms.DialogResult.Yes)
                                SomeMethodsClass.SetBackDockForCommit(fileNum, abits);
                            }

                            //если заявлений нет, то это чисто новое заявление
                            //но даже если они есть, то пусть догружают
                            CardFromInet crd = new CardFromInet(null, fileNum, false);
                            crd.ToUpdateList += new UpdateListHandler(UpdateDataGrid);
                            crd.Show();
                        }
                    }
                    catch (Exception exc)
                    {
                        WinFormsServ.Error(exc.Message);
                        tbPersonNum.Text = "";
                        tbPersonNum.Focus();
                    }
                }
                else
                {
                    //если заявления по данному номеру коммита в базе уже есть, то нужно открыть карточку (т.к. коммит в онлайне ещё жив)
                    {
                        using (PriemEntities context = new PriemEntities())
                        {
                            extAbit abit = (from ab in context.extAbit
                                            where ab.CommitNumber == fileNum
                                            select ab).FirstOrDefault();

                            if (abit != null)
                            {
                                string fio = abit.FIO;
                                string num = abit.PersonNum;
                                string persId = abit.PersonId.ToString();

                                var dr = MessageBox.Show(string.Format("Абитуриент {0} с данным штрих-кодом уже импортирован в базу.\nОткрыть карточку абитуриента?", fio), "Внимание", MessageBoxButtons.YesNo);
                                if (dr == System.Windows.Forms.DialogResult.Yes)
                                    MainClass.OpenCardPerson(persId, null, null);
                            }
                        }
                    }

                    //UpdateDataGrid();
                    //using (PriemEntities context = new PriemEntities())
                    //{
                    //    extAbit abit = (from ab in context.extAbit
                    //                    where ab.CommitNumber == fileNum
                    //                    select ab).FirstOrDefault();

                    //    string fio = abit.FIO;
                    //    string num = abit.PersonNum;
                    //    string persId = abit.PersonId.ToString();

                    //    WinFormsServ.Search(this.dgvAbiturients, "PersonNum", num);
                    //    DialogResult dr = MessageBox.Show(string.Format("Абитуриент {0} с данным номером баркода уже импортирован в базу.\nОткрыть карточку абитуриента?", fio), "Внимание", MessageBoxButtons.YesNo);
                    //    if (dr == System.Windows.Forms.DialogResult.Yes)
                    //        MainClass.OpenCardPerson(persId, this, null);
                    //}
                }
            }
            tbPersonNum.Text = "";
            tbPersonNum.Focus();
            loadClass.CloseDB();  
        }

        private void PersonList_Load(object sender, EventArgs e)
        {
            tbPersonNum.Focus();
        }

        private void PersonList_Activated(object sender, EventArgs e)
        {
            tbPersonNum.Focus();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            UpdateDataGrid();            
        }

        private void tbNumber_TextChanged(object sender, EventArgs e)
        {
            WinFormsServ.Search(this.dgvAbiturients, "PersonNum", tbNumber.Text);
        }

        private void chbShowAll_CheckedChanged(object sender, EventArgs e)
        {
            UpdateDataGrid();
        }

        protected override void btnRemove_Click(object sender, EventArgs e)
        {
            if (MainClass.IsPasha())
            {
                if (MessageBox.Show("Удалить записи?", "Удаление", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    foreach (DataGridViewRow dgvr in Dgv.SelectedRows)
                    {
                        string itemId = dgvr.Cells["Id"].Value.ToString();
                        Guid g = Guid.Empty;
                        Guid.TryParse(itemId, out g);
                        try
                        {
                            using (PriemEntities context = new PriemEntities())
                            {
                                context.Person_deleteAllInfo(g);
                            }
                        }
                        catch (Exception ex)
                        {
                            WinFormsServ.Error("Ошибка удаления данных" + ex.Message);
                            //goto Next;
                        }
                    //Next: ;
                    }
                    MainClass.DataRefresh();
                }
            }
        }
    }
}