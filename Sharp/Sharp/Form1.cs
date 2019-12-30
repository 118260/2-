using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Sharp
{
    public partial class Laba : Form
    {
        int TriggerFlag;
        public Laba()
        {
            InitializeComponent();

        }





        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Заводим активную Excel книгу и страницу
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;

            //Открытие файла
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Выберите файл";
            ofd.Filter = "Файл Excel|*.XLSX;*.XLS | Все файлы(*.*)|*.*";
            ofd.ShowDialog();
            System.Data.DataTable tb = new System.Data.DataTable();
            string filename = ofd.FileName;
            ExcelWorkBook = ExcelApp.Workbooks.Open(ofd.FileName);

            //Создание первой активной страницы
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);



            //Цикл создания колонок DatagridView
            for (int a = 1; a < 5; a++)
            {

                String Colomn = "Colomn";
                String NumberColomn = (a - 1).ToString();
                String ColomnName = Colomn + NumberColomn;

                dataGridView1.Columns.Add(ColomnName, null);
                dataGridView1.Columns[a - 1].HeaderText = ExcelWorkSheet.Cells[1, a].value;
            }



            //Цикл добавление строк и копирования из Excel, с конвертацией типов в String, чтобы не вылезали эксепшены в нескольких местах
            //да и в целом выглядело опрятнее

            for (int i = 2; ExcelWorkSheet.Cells[i, 4].value != null; i++)
            {
                dataGridView1.Rows.Add(null, null, null, null);
                for (int j = 1; j < 5; j++)
                {

                    dataGridView1.Rows[i - 2].Cells[j - 1].Value = ExcelWorkSheet.Cells[i, j].value;
                    String temp = Convert.ToString(dataGridView1.Rows[i - 2].Cells[j - 1].Value);
                    dataGridView1.Rows[i - 2].Cells[j - 1].Value = temp;


                }


            }


           //Цикл вывода просрока через Message Box
            
            for (int index = 3; dataGridView1.Rows[index - 2].Cells[3].Value != null; index++)
            {

                /*Вот тут я мог бы оперировать типами DateTime настоящего времени и даты просрока, но для вычисления разности не работал TimeSpan, 
                поэтому вот так вот*/

                //=====================================================
                // Берем тип DateTime даты нашего просрока, конвертируем в String,
                //убираем лишнее в формате String и конвертируем в int для дальнейших вычислений
                String FullDate = Convert.ToString(dataGridView1.Rows[index - 2].Cells[3].Value); 
                    FullDate = FullDate.Remove(10, 8);

                    String SDay = FullDate.Remove(2, 8);
                    int Day = Convert.ToInt32(SDay);

                    String MonthYear = FullDate.Substring(3);
                    String SMonth = MonthYear.Remove(2, 5);
                    int Month = Convert.ToInt32(SMonth);


                    String SYear = FullDate.Substring(6);
                    int Year = Convert.ToInt32(SYear);


                //=====================================================
                //!!!!! Берем тип DateTime настоящего времени, конвертируем в String, убираем лишнее в формате String
                //И конвертируем в int для дальнейших вычислений

                DateTime DTNow = DateTime.Now;
                    String Now = Convert.ToString(DTNow);
                    Now = Now.Remove(10, 9);


                    String SNDay = Now.Remove(2, 8);
                    int NDay = Convert.ToInt32(SNDay);

                    String SNMonthYear = Now.Substring(3);
                    String SNMonth = SNMonthYear.Remove(2, 5);
                    int NMonth = Convert.ToInt32(SNMonth);


                    String SNYear = Now.Substring(6);
                    int NYear = Convert.ToInt32(SNYear);

                //Здесь мы просто проверяем на просрок через условие и ставим тригеру 0/1 в зависимости просрочилось ли
                // 1 - просрок, 0 - будет ещё лежать
                    if ((NYear == Year) && (NMonth == Month) && (NDay == Day))
                    {
                        TriggerFlag = 1; 
                    }

                    else if ((NYear == Year) && (NMonth == Month) && (NDay != Day))
                    {
                        if (NDay > Day)
                        {
                            TriggerFlag = 1;
                        }
                        else
                        {
                            TriggerFlag = 0;
                        }
                    }
                    else if ((NYear == Year) && (NMonth != Month))
                    {
                        if (NMonth > Month)
                        {
                            TriggerFlag = 1;
                        }
                        else
                        {
                            TriggerFlag = 0;
                        }
                    }
                    else if (NYear != Year)
                    {
                        if (NYear > Year)
                        {
                            TriggerFlag = 1;
                        }
                        else
                        {
                            TriggerFlag = 0;
                        }
                    }

                    // Тут мы проверяем, если флаг равен 1, то он выводит пушик через Massagebox о том, какой товар просрочился

                    if (TriggerFlag == 1)
                    {
                        String ID = dataGridView1.Rows[index - 2].Cells[0].Value.ToString();
                        String VendorCode = dataGridView1.Rows[index - 2].Cells[1].Value.ToString();
                        String NameProduct = dataGridView1.Rows[index - 2].Cells[2].Value.ToString();
                        String FullName = "ID:" + ID + " " + "Артикул:" + VendorCode + " " + "Наименование:" + NameProduct;
                        MessageBox.Show(FullName, "ПРОСРОЧЕНО");

                    }
                 
    }
              
            //Закрытие Excel, чтобы не жрало память
            ExcelApp.Quit();
  }
        

        private void Form1_Load(object sender, EventArgs e)
        {

        }

       
    }
}

