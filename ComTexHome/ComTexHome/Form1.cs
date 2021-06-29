using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace ComTexHome
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
            button2.Visible = false;
            comboBox1.Visible = false;
            dateTimePicker1.Visible = false;
            dateTimePicker2.Visible = false;
            label1.Visible = label2.Visible = label3.Visible = label4.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application E = new Excel.Application();
            E.Visible = true;
            E.Workbooks.Open(@"C:\Users\Acer\Desktop\ComTexHome\ComTexHome\bin\Debug\test.xlsx");//Указать путь к файлу
            var Sh = E.ActiveSheet;
            string[] arr = new string[5];
            dataGridView1.ColumnCount = 5;
            dataGridView1.Rows.Clear();
            int r = 0;
            do
            {
                r++;
                arr[0] = Sh.Cells[r + 1, 1].Text;
                if (arr[0] == "") break;
                arr[1] = Sh.Cells[r + 1, 2].Text;
                arr[2] = Sh.Cells[r + 1, 3].Text;
                arr[3] = Sh.Cells[r + 1, 4].Text;
                arr[4] = Sh.Cells[r + 1, 5].Text;
                dataGridView1.Rows.Add(arr[0], arr[1], arr[2], arr[3], arr[4]);
            } while (true);
            E.Quit(); 
            E = null;
            button2.Visible = true;
            button1.Visible = false;
            dateTimePicker1.Visible = true;
            dateTimePicker2.Visible = true;
            comboBox1.Visible = true;
            label1.Visible = label2.Visible = label3.Visible = label4.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var W = new Word.Application();
            W.Visible = false;
            var D = W.Documents.Add();
            var t = W.Selection;
            t.Font.Bold = 0;
            t.Font.Size = 12;

            object obj = 0;
            object obj1 = 0;
            Word.Range tableLocation = D.Range(ref obj, ref obj1);
            D.Tables.Add(t.Range, 21, 5);
            Word.Table newTable = D.Tables[1];
            for (int i = 0; i < 5; i++)
            {
                newTable.Cell(1, i + 1).Range.Text = dataGridView1.Columns[i].HeaderText;
                for (int j = 0; j < 20; j++)
                {
                    newTable.Cell(j + 2, i + 1).Range.Text = dataGridView1.Rows[j].Cells[i].Value.ToString();
                }
            }
            newTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDouble;
            newTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDouble;
            newTable.Range.Paragraphs.Last.Next().Range.Select();

            //3
            t.TypeParagraph();
            t.TypeText("3. Список номеров вагонов, использовавшихся в первом полугодии прошлого года.\n");
            List<string> L = new List<string>();
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                if (Convert.ToDateTime(dataGridView1.Rows[i].Cells[3].Value) < Convert.ToDateTime("01.07.2019"))
                {
                    if (L.IndexOf(dataGridView1.Rows[i].Cells[1].Value.ToString()) == -1)
                        L.Add(dataGridView1.Rows[i].Cells[1].Value.ToString());
                }
            }
            foreach (string y in L)
            {
                t.TypeText(y + ", ");
            }

            //4
            t.TypeParagraph();
            t.TypeText("4. Средняя стоимость перевозок по каждому из встречающихся грузов.\n");
            string[] Name = new string[0];
            int[] Count = new int[0];
            int[] Cost = new int[0];
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                bool tf = false;
                int when = 0;
                for (int j = 0; j < Name.Length; j++)
                {
                    if (Name[j] == dataGridView1.Rows[i].Cells[0].Value.ToString())
                    {
                        tf = true; when = j;
                    }
                }
                if (tf)
                {
                    Count[when]++;
                    Cost[when] += Convert.ToInt32(dataGridView1.Rows[i].Cells[2].Value);
                }
                else
                {
                    Array.Resize(ref Name, Name.Length + 1);
                    Name[Name.Length - 1] = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    Array.Resize(ref Count, Count.Length + 1);
                    Count[Count.Length - 1] = 1;
                    Array.Resize(ref Cost, Cost.Length + 1);
                    Cost[Cost.Length - 1] = Convert.ToInt32(dataGridView1.Rows[i].Cells[2].Value);
                }
            }

            var table = D.Tables.Add(t.Range, Name.Count() + 1, 2);
            table.Cell(1, 1).Range.Text = "Груз";
            table.Cell(1, 2).Range.Text = "Стоимость";
            for (int i = 0; i < Name.Length; i++)
            {
                table.Cell(i + 2, 1).Range.Text = Name[i];
                table.Cell(i + 2, 2).Range.Text = Convert.ToString(Cost[i] / Count[i]);
            }
            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDouble;
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDouble;
            table.Range.Paragraphs.Last.Next().Range.Select();


            //5
            t.TypeParagraph();
            t.TypeText(String.Format("5. Количество дней использования каждого из вагонов в месяц {0} текущего года.\n", comboBox1.Items[comboBox1.SelectedIndex]));
            string[] Num = new string[0];
            int[] Days = new int[0];
            DateTime first = new DateTime(2020, comboBox1.SelectedIndex + 1, 1);
            DateTime last = new DateTime(2020, comboBox1.SelectedIndex + 2, 1).AddDays(-1);
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                bool tf = false;
                int when = 0;
                for (int j = 0; j < Num.Length; j++)
                {
                    if (Num[j] == dataGridView1.Rows[i].Cells[1].Value.ToString())
                    {
                        tf = true; when = j;
                    }
                }
                if (!tf)
                {
                    Array.Resize(ref Num, Num.Length + 1);
                    Num[Num.Length - 1] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                    Array.Resize(ref Days, Days.Length + 1);
                    Days[Days.Length - 1] = 0;
                    when = Days.Length - 1;
                }
                DateTime start = Convert.ToDateTime(dataGridView1.Rows[i].Cells[3].Value);
                DateTime end = Convert.ToDateTime(dataGridView1.Rows[i].Cells[4].Value);

                if (start >= first && start <= last)
                {
                    if (end >= first && end <= last)
                        Days[when] += (int)end.Subtract(start).TotalDays + 1;
                    else
                        Days[when] += (int)last.Subtract(start).TotalDays + 1;
                }
                else
                {
                    if (start < first)
                    {
                        if (end >= first && end <= last)
                            Days[when] += (int)end.Subtract(first).TotalDays + 1;
                        else
                            if (end > last)
                            Days[when] += (int)last.Subtract(first).TotalDays + 1;
                    }
                }
            }

            var table3 = D.Tables.Add(t.Range, Num.Count() + 1, 2);
            table3.Cell(1, 1).Range.Text = "Номер вагона";
            table3.Cell(1, 2).Range.Text = "Количество дней";
            for (int i = 0; i < Num.Length; i++)
            {
                table3.Cell(i + 2, 1).Range.Text = Num[i];
                table3.Cell(i + 2, 2).Range.Text = Days[i].ToString();
            }
            table3.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDouble;
            table3.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDouble;
            table3.Range.Paragraphs.Last.Next().Range.Select();

            //6
            first = dateTimePicker1.Value;
            last = dateTimePicker2.Value;
            int Sum = 0;
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                DateTime start = Convert.ToDateTime(dataGridView1.Rows[i].Cells[3].Value);
                DateTime end = Convert.ToDateTime(dataGridView1.Rows[i].Cells[4].Value);

                if ((start >= first && start <= last) || (end >= first && end <= last))
                {
                    Sum += Convert.ToInt32(dataGridView1.Rows[i].Cells[2].Value);
                }
                else
                {
                    if (start < first && end > last)
                    {
                        Sum += Convert.ToInt32(dataGridView1.Rows[i].Cells[2].Value);
                    }
                }
            }
            t.TypeParagraph();
            t.TypeText(String.Format("6. Общая стоимость перевозок за период {0} - {1} равна {2}", first.ToShortDateString(), last.ToShortDateString(), Sum));

            //t.Font.Bold = -1;
            //t.Font.Size = 24;

            //Диаграмма
            var E2 = new Excel.Application();
            E2.Visible = true;
            E2.Workbooks.Open(@"C:\Users\Acer\Desktop\ComTexHome\ComTexHome\bin\Debug\test.xlsx");//Указать путь к файлу
            var Sh = E2.ActiveSheet;
            E2.Sheets.Add();
            var Sh2 = E2.Worksheets[1];
            Sh2.Name = "Диаграмма";

            Sh2.Cells[1, 1] = "Груз";
            Sh2.Cells[1, 2] = "Стоимость";
            for (int i = 0; i < Name.Length; i++)
            {
                Sh2.Cells[i + 2, 1] = Name[i];
                Sh2.Cells[i + 2, 2] = Cost[i] / Count[i];
            }
            Excel.Range rng = Sh2.Range(Sh2.Cells[1, 1], Sh2.Cells[Name.Length, 2]);
            var Ch = E2.Charts.Add();
            Ch.Location(Excel.XlChartLocation.xlLocationAsObject, "Диаграмма");
            Ch = E2.ActiveChart;
            Ch.ChartTitle.Text = "Средняя стоимость перевозки каждого груза";
            Ch.HasLegend = false;

            Ch.ChartArea.Select();
            Ch.ChartArea.Copy();
            t.TypeParagraph();
            t.Paste();
            string s = String.Format(@"C:\Users\Acer\Desktop\ComTexHome\ComTexHome\bin\Debug\test - Абдуназаров Достон {0}   {1}-{2}-{3}.docx", DateTime.Now.ToShortDateString(), DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
            D.SaveAs(s);
            W.Quit(); W = null;
            E2.Quit(); E2 = null;
            button1.Visible = false;
        }
    }
}
