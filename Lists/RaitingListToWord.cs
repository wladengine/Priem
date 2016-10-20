using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using EducServLib;
using BDClassLib;
using WordOut;

using RtfWriter;
using PriemLib;
using System.Data.Entity.Core.Objects;
using System.Threading.Tasks;


namespace Priem 
{
    public partial class RaitingListToWord : Form
    {
        bool _hasOlymp;
        DataGridView dgvAbits;

        public RaitingListToWord(bool HasOlymp, DataGridView _d)
        {
            InitializeComponent();
            _hasOlymp = HasOlymp;
            dgvAbits = _d;
            FillGrid();
        }
        private void FillGrid()
        {
            List<PriorityToWordItem> lst = new List<PriorityToWordItem>();
            lst.Add(new PriorityToWordItem() { IsPrint = true, Name = "№ п/п", Width = 8.5m });
            lst.Add(new PriorityToWordItem() { IsPrint = true, Name = "Рег.номер", Width = 18.5m });
            lst.Add(new PriorityToWordItem() { IsPrint = true, Name = "ФИО", Width = 40.0m });
            lst.Add(new PriorityToWordItem() { IsPrint = true, Name = "Сумма баллов", Width = 15.0m });
            lst.Add(new PriorityToWordItem() { IsPrint = true, Name = "Проф. экзамен", Width = 15.0m });
            lst.Add(new PriorityToWordItem() { IsPrint = true, Name = "Доп. экзамен", Width = 15.0m });
            lst.Add(new PriorityToWordItem() { IsPrint = true, Name = "Конкурс", Width = 15.0m });
            lst.Add(new PriorityToWordItem() { IsPrint = true, Name = "Подлинники", Width = 18.0m });
            lst.Add(new PriorityToWordItem() { IsPrint = true, Name = "Контакты", Width = 45.0m });
            lst.Add(new PriorityToWordItem() { IsPrint = true, Name = "Медалист", Width = 15.5m });
            lst.Add(new PriorityToWordItem() { IsPrint = true, Name = "Серия док. об обр.", Width = 15.1m });
            lst.Add(new PriorityToWordItem() { IsPrint = true, Name = "ср. балл", Width = 15.1m });
            lst.Add(new PriorityToWordItem() { IsPrint = true, Name = "Ретинг. коэфф.", Width = 15.1m });
            if (_hasOlymp)
                lst.Add(new PriorityToWordItem() { IsPrint = true, Name = "Олимпиада", Width = 15.1m });

            DataGridViewCheckBoxCell cell1 = new DataGridViewCheckBoxCell();
            DataGridViewCheckBoxColumn chb_col = new DataGridViewCheckBoxColumn();
            chb_col.CellTemplate = cell1;
            chb_col.Name = "На_печать";
            chb_col.HeaderText = "Печать";
            dgv.Columns.Add(chb_col);

            dgv.Columns.Add("Столбец", "Столбец");
            dgv.Columns.Add("Ширина", "Ширина");

            foreach (var s in lst)
            {
                dgv.Rows.Add();
                dgv.Rows[dgv.Rows.Count - 1].Cells[0].Value = s.IsPrint;
                dgv.Rows[dgv.Rows.Count - 1].Cells[1].Value = s.Name;
                dgv.Rows[dgv.Rows.Count - 1].Cells[2].Value = s.Width;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            List<PriorityToWordItem> lst = new List<PriorityToWordItem>();
            foreach (DataGridViewRow rw in dgv.Rows)
            {
                lst.Add(new PriorityToWordItem() { IsPrint = (bool)rw.Cells[0].Value, Name = rw.Cells[1].Value.ToString(), Width = (decimal)rw.Cells[2].Value });
            }
            int addcols = (int)num.Value;
            for (int i = 0; i < addcols; i++)
            {
                lst.Add(new PriorityToWordItem() { IsPrint = true, Name = "", Width = 15.1m });
            }

            ToWord(lst);
        }
        private void ToWord(List<PriorityToWordItem> lst)
        {
            int rowCount = dgvAbits.Rows.Count;
            if (rowCount == 0)
                return;
            try
            {
                float margin = (float)(20.0m * RtfConstants.MILLIMETERS_TO_POINTS);
                RtfDocument doc = new RtfDocument(PaperSize.A4, PaperOrientation.Landscape, Lcid.Russian, margin, margin, margin, margin);

                RtfTable table = doc.addTable(rowCount + 1, lst.Where(x => x.IsPrint).Count(), (float)(276.1m * RtfConstants.MILLIMETERS_TO_POINTS));

                // Устанавливаем ширину столбцов таблицы (в миллиметрах)
                int i = 0;
                foreach (var s in lst.Where(x => x.IsPrint))
                {
                    
                    table.setColWidth(i, (float)(s.Width * RtfConstants.MILLIMETERS_TO_POINTS));
                    i++;
                }
                i = 0;
                foreach (var s in lst.Where(x => x.IsPrint))
                {
                    table.cell(0, i).addParagraph().Text = s.Name; 
                    i++;
                }

                for (int j = 0; j < lst.Where(x => x.IsPrint).Count(); j++)
                {
                    // Устанавливаем горизонтальное и вертикальное выравнивание текста "по центру" в каждой ячейке таблицы
                    table.cell(0, j).Alignment = Align.Center;
                    table.cell(0, j).AlignmentVertical = AlignVertical.Middle;
                }

                int r = 0;
                foreach (DataGridViewRow row in dgvAbits.Rows)
                {
                    ++r;
                    i = 0;
                    if (lst[0].IsPrint)
                    {
                        table.cell(r, i).addParagraph().Text = r.ToString();
                        i++;
                    }
                    if (lst[1].IsPrint)
                    {
                        table.cell(r, i).addParagraph().Text = row.Cells["Рег_Номер"].Value.ToString();
                        i++;
                    }
                    if (lst[2].IsPrint)
                    {
                        table.cell(r, i).addParagraph().Text = row.Cells["ФИО"].Value.ToString();
                        i++;
                    }
                    if (lst[3].IsPrint)
                    {
                        table.cell(r, i).addParagraph().Text = row.Cells["Сумма баллов"].Value.ToString();
                        i++;
                    }
                    if (lst[4].IsPrint)
                    {
                        table.cell(r, i).addParagraph().Text = row.Cells["Проф. экзамен"].Value.ToString();
                        i++;
                    }
                    if (lst[5].IsPrint)
                    {
                        table.cell(r, i).addParagraph().Text = row.Cells["Доп. экзамен"].Value.ToString();
                        i++;
                    }
                    if (lst[6].IsPrint)
                    {
                        table.cell(r, i).addParagraph().Text = row.Cells["Конкурс"].Value.ToString();
                        i++;
                    }
                    if (lst[7].IsPrint)
                    {
                        table.cell(r, i).addParagraph().Text = row.Cells["Подлинники документов"].Value.ToString();
                        i++;
                    }
                    if (lst[8].IsPrint)
                    {
                        table.cell(r, i).addParagraph().Text = row.Cells["Контакты"].Value.ToString();
                        i++;
                    }
                    if (lst[9].IsPrint)
                    {
                        table.cell(r, i).addParagraph().Text = row.Cells["Медалист"].Value.ToString();
                        i++;
                    }
                    if (lst[10].IsPrint)
                    {
                        table.cell(r, i).addParagraph().Text = MainClass.dbType == PriemType.PriemMag ? row.Cells["Серия диплома"].Value.ToString() : row.Cells["Серия аттестата"].Value.ToString();
                        i++;
                    }
                    if (lst[11].IsPrint)
                    {
                        table.cell(r, i).addParagraph().Text = row.Cells["Средний балл"].Value.ToString();
                        i++;
                    }
                    if (lst[12].IsPrint) 
                    {
                        table.cell(r, i).addParagraph().Text = row.Cells["Рейтинговый коэффициент"].Value.ToString(); 
                        i++;
                    }

                    if (_hasOlymp && lst[13].IsPrint)
                    { 
                        table.cell(r, i).addParagraph().Text = row.Cells["Олимпиада"].Value.ToString();
                        i++;
                    }


                    for (int j = 0; j < lst.Where(x => x.IsPrint).Count(); j++)
                    {
                        // Устанавливаем горизонтальное и вертикальное выравнивание текста "по центру" в каждой ячейке таблицы
                        table.cell(r, j).Alignment = Align.Center;
                        table.cell(r, j).AlignmentVertical = AlignVertical.Middle;
                    }
                }

                // Задаём толщину внутренних границ таблицы
                table.setInnerBorder(RtfWriter.BorderStyle.Single, 0.5f);
                // Задаём толщину внешних границ таблицы
                table.setOuterBorder(RtfWriter.BorderStyle.Single, 0.5f);

                doc.save(MainClass.saveTempFolder + "\\RatingList.rtf");

                // ==========================================================================
                // Открываем сохранённый RTF файл
                // ==========================================================================
                WordDoc wd = new WordDoc(string.Format(@"{0}\RatingList.rtf", MainClass.saveTempFolder));
            }
            catch (Exception ex)
            {
                WinFormsServ.Error("Ошибка при составлении списка:\n" + ex.Message +
                    ex.InnerException == null ? "" : ("\nВнутреннее исключение:\n" + ex.InnerException.Message));
            }
        }

    }


    public class PriorityToWordItem
    {
        public bool IsPrint;
        public string Name;
        public decimal Width;

    }
}
