using System;
using System.Data;
using System.Linq;
using System.Data.Common;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Data.SQLite;
using Microsoft.Office;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;

namespace Curator_v_0_5
{
    
    public partial class Form1 : Form
       

    {
        public string dbName = "curator.db"; 
        public SQLiteConnection con;
        public SQLiteCommand cmd;
        public System.Data.DataTable dt;
        public SQLiteDataAdapter adapter;

        public class student
        {
            /*
            * оголошуємо поля класу
            */

            public string prizvishe; // Прізвище
            public string name; //Ім'я
            public string po_batk; //По батькові
            public string data_ok; //Дата народження 
            public string phone; // Телефон
            public string adresa; //Адреса проживання
            public int starost; // Чи є старостою
            public string forma; // Форма навчання
            public string stat; // Стать
            public string grupa; // Група

            public void set_stud(string a1, string a2, string a3, string a4, 
                                    string a5, string a6, int a7, string a8, string a9, string a10)
            {
                prizvishe = a1;
                name = a2;
                po_batk = a3;
                data_ok = a4;
                phone = a5;
                adresa= a6;
                starost = a7;
                forma = a8;
                stat = a9;
                grupa = a10;
        }


        }


        public Form1()
        {
            InitializeComponent();
            dt = new System.Data.DataTable(); // Создаем объект DataTable
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме
            con = new SQLiteConnection(); //создание нового подключения
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";
            adapter = new SQLiteDataAdapter("Select * from Студент ORDER BY Фамилия ASC", con);
            SQLiteCommandBuilder builder = new SQLiteCommandBuilder(adapter);
            try
            {
                adapter.UpdateCommand = builder.GetUpdateCommand();
                adapter.InsertCommand = builder.GetInsertCommand();
                adapter.DeleteCommand = builder.GetDeleteCommand();
                adapter.Fill(dt);
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        public void Export_Data_To_Word(DataGridView DGV, string filename)
        {

            if (DGV.Rows.Count != 0)
            {
                int RowCount = DGV.Rows.Count;
                int ColumnCount = DGV.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                //add rows
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                    } //end row loop
                } //end column loop

                Word.Document oDoc = new Word.Document();
                oDoc.Application.Visible = true;

                //page orintation
                oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;


                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";

                    }
                }

                //table format
                oRange.Text = oTemp;
                object oMissing = System.Reflection.Missing.Value;
                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();

                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();

                //header row style
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Tahoma";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 14;

                //add header row manually
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
                }

                //table style 
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //header text
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.Text = "your header text";
                    headerRange.Font.Size = 16;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }

                //save the file

                oDoc.SaveAs(filename, ref oMissing, ref oMissing, ref oMissing,
    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
    ref oMissing, ref oMissing);

                //NASSIM LOUCHANI
            }
        }


        //выбор файла базы данных
        private void button1_Click(object sender, EventArgs e)
        {
            // выбор файла базы данных
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Clear();
                dbName = openFileDialog1.FileName;

                textBox1.Text = dbName;
                MessageBox.Show("Файл бази даних вибрано.");

            }
        }
        //вывод и редактирование таблицы "Студенты"
        private void button2_Click(object sender, EventArgs e)
        {
            student[] st = new student[50];  // створення масиву об'єкта класа student

            dt = new System.Data.DataTable(); // Создаем объект DataTable
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме
            con = new SQLiteConnection(); //создание нового подключения
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";
            adapter = new SQLiteDataAdapter("Select * from Студент ORDER BY Фамилия ASC", con);

            //adapter = new SQLiteDataAdapter("Select * from Студент ORDER BY Фамилия ASC", con);
            SQLiteCommandBuilder builder = new SQLiteCommandBuilder(adapter);
            try
            {
                adapter.UpdateCommand = builder.GetUpdateCommand();
                adapter.InsertCommand = builder.GetInsertCommand();
                adapter.DeleteCommand = builder.GetDeleteCommand();
                adapter.Fill(dt);
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.Message);
            }
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Прізвище";
            dataGridView1.Columns[2].HeaderText = "Ім'я";
            dataGridView1.Columns[3].HeaderText = "По батькові";
            dataGridView1.Columns[4].HeaderText = "Дата народження";
            dataGridView1.Columns[5].HeaderText = "Телефон";
            dataGridView1.Columns[6].HeaderText = "Адреса проживання";
            dataGridView1.Columns[7].HeaderText = "Відомості про батьків";
            dataGridView1.Columns[8].HeaderText = "Додаткова інформація";
            dataGridView1.Columns[9].HeaderText = "Чи є старостою";
            dataGridView1.Columns[10].HeaderText = "Форма навчання";
            dataGridView1.Columns[11].HeaderText = "Стать";
            dataGridView1.Columns[12].HeaderText = "Група";
          //  dataGridView1.Columns[13].HeaderText = "Додаткова інформація";


        }
        //вывод и редактирование таблицы "Дисциплины"
        private void button3_Click(object sender, EventArgs e)
        {
            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме
            con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";
            adapter = new SQLiteDataAdapter("Select * from Дисциплина  ORDER BY Код_дисциплины ASC", con);
            SQLiteCommandBuilder builder = new SQLiteCommandBuilder(adapter);
            try
            {
                adapter.UpdateCommand = builder.GetUpdateCommand();
                adapter.InsertCommand = builder.GetInsertCommand();
                adapter.DeleteCommand = builder.GetDeleteCommand();
                adapter.Fill(dt);
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.Message);
            }

            dataGridView1.Columns[0].HeaderText = "Код дисципліни";
            dataGridView1.Columns[1].HeaderText = "Найменування";
            dataGridView1.Columns[2].HeaderText = "Викладач";
            



        }
        
        private void button4_Click(object sender, EventArgs e)
        {
            con = new SQLiteConnection();
            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt;          
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";
            adapter = new SQLiteDataAdapter("Select * FROM Расписание ORDER BY Дата ASC, Номер_урока ASC", con);
            SQLiteCommandBuilder builder = new SQLiteCommandBuilder(adapter);
            try
            {
                adapter.UpdateCommand = builder.GetUpdateCommand();
                adapter.InsertCommand = builder.GetInsertCommand();
                adapter.DeleteCommand = builder.GetDeleteCommand();
                adapter.Fill(dt);
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.Message);
            }

            dataGridView1.Columns[0].HeaderText = "Дата заняття";
            dataGridView1.Columns[1].HeaderText = "Номер пари";
            dataGridView1.Columns[2].HeaderText = "Дисципліна";
          
        }

       
        private void button5_Click(object sender, EventArgs e)
        {
            con = new SQLiteConnection();
            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt;        
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";
            adapter = new SQLiteDataAdapter("Select * from Посещаемость ORDER BY Дата ASC, Номер_урока ASC ", con);
            SQLiteCommandBuilder builder = new SQLiteCommandBuilder(adapter);
            try
            {
            adapter.UpdateCommand = builder.GetUpdateCommand();
            adapter.InsertCommand = builder.GetInsertCommand();
            adapter.DeleteCommand = builder.GetDeleteCommand();
            adapter.Fill(dt);
            }
            catch (SQLiteException ex)
            {
            MessageBox.Show(ex.Message);
            }

           /* dataGridView1.Columns[1].HeaderText = "Номер лекції";
            dataGridView1.Columns[2].HeaderText = "Шифр студента";
            dataGridView1.Columns[3].HeaderText = "Відмітка";*/

        }

        
        private void button6_Click(object sender, EventArgs e)
        {
           
            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt;         
            con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";
            adapter = new SQLiteDataAdapter("Select * from Успеваемость ORDER BY Шифр_студента ASC , Код_дисциплины ASC", con);
            SQLiteCommandBuilder builder = new SQLiteCommandBuilder(adapter);
            try
            {
               adapter.UpdateCommand = builder.GetUpdateCommand();
               adapter.InsertCommand = builder.GetInsertCommand();
               adapter.DeleteCommand = builder.GetDeleteCommand();
               adapter.Fill(dt);
        
            }
             catch (SQLiteException ex)
            {
               MessageBox.Show(ex.Message);
            }

            dataGridView1.Columns[0].HeaderText = "Шифр студента";
            dataGridView1.Columns[1].HeaderText = "Код дисципліни";
            dataGridView1.Columns[2].HeaderText = "Атестація";

        }

        
        private void button7_Click(object sender, EventArgs e)
        {
            
            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt; 
            con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";
            adapter = new SQLiteDataAdapter("Select * from Задания", con);
            SQLiteCommandBuilder builder = new SQLiteCommandBuilder(adapter);
            try
            {
                adapter.UpdateCommand = builder.GetUpdateCommand();
                adapter.InsertCommand = builder.GetInsertCommand();
                adapter.DeleteCommand = builder.GetDeleteCommand();
                adapter.Fill(dt);

            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.Message);
            }


            dataGridView1.Columns[0].HeaderText = "Код завдання";
            dataGridView1.Columns[1].HeaderText = "Код дисципліни";
            dataGridView1.Columns[2].HeaderText = "Вид завдання";
            dataGridView1.Columns[3].HeaderText = "Дата здачі";

        }
        //вывод и редактирование таблицы "Темы"
        private void button8_Click(object sender, EventArgs e)
        {
            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме           
            con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";
            adapter = new SQLiteDataAdapter("Select * from Темы", con);
            SQLiteCommandBuilder builder = new SQLiteCommandBuilder(adapter);
            try
            {
                adapter.UpdateCommand = builder.GetUpdateCommand();
                adapter.InsertCommand = builder.GetInsertCommand();
                adapter.DeleteCommand = builder.GetDeleteCommand();
                adapter.Fill(dt);

            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.Message);
            }

            dataGridView1.Columns[0].HeaderText = "Код завдання";
            dataGridView1.Columns[1].HeaderText = "Шифр студента";
            dataGridView1.Columns[2].HeaderText = "Відмітка про здачу";
            dataGridView1.Columns[3].HeaderText = "Тема завдання";
        }
        //обновление данных в базе данных
        private void button9_Click(object sender, EventArgs e)
        {
            adapter.Update(dt);//обновляем данные в базе данных
        }
        //экспорт таблицы в Excel
        private void button10_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook ExcelWorkBook;
            Worksheet ExcelWorkSheet;
            //Создаем книгу Excel
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            //Создаем и заполняем таблицу (Лист Excel)
            ExcelWorkSheet = (Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            
            //жирный шрифт для первой строки
            ExcelApp.Rows[1].Font.Bold = true;
            //заполнение наименований столбцов
            for (int j = 0; j < dataGridView1.ColumnCount; j++)
            {
                ExcelApp.Cells[1, j + 1] = dataGridView1.Columns[j].HeaderText;
            }
            //заполнение таблицы
            
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            //Изменение размера ячеек по содержимому
            ExcelWorkSheet.Columns.AutoFit();
            ExcelWorkSheet.Rows.AutoFit();
            //Вызов созданной книги в Excel
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;

        }
        //Просмотр отчетов
        
        private void button11_Click(object sender, EventArgs e)
        {
            con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";

            cmd = new SQLiteCommand();
            cmd.Connection = con;

            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме

            con.Open(); // открываем соединение с БД

            cmd.CommandText = "SELECT р.Дата, р.Номер_урока,  д.Наименование, д.Преподаватель FROM Расписание р INNER JOIN Дисциплина д ON(р.Код_дисциплины = д.Код_дисциплины) ORDER BY Дата ASC, Номер_урока ASC";
            dt.Load(cmd.ExecuteReader()); // выполняем SQL-запрос

            con.Close(); // закрываем соединение с БД


            dataGridView1.Columns[0].HeaderText = "Дата заняття";
            dataGridView1.Columns[1].HeaderText = "Номер пари";
            dataGridView1.Columns[2].HeaderText = "Дисципліна";
            dataGridView1.Columns[3].HeaderText = "Викладач";

        }
        //Отчет по пропускам
        private void button12_Click(object sender, EventArgs e)
        {
            con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";

            cmd = new SQLiteCommand();
            cmd.Connection = con;

            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме

            con.Open(); // открываем соединение с БД
            cmd.CommandText = "SELECT Пропуск AS Причина, count(*) AS Количество FROM Посещаемость GROUP BY Пропуск";
            dt.Load(cmd.ExecuteReader()); // выполняем SQL-запрос
            con.Close(); // закрываем соединение с БД

            dataGridView1.Columns[0].HeaderText = "Відмітка";
            dataGridView1.Columns[1].HeaderText = "Кількість";
        }
        //Пропуски по дисциплине
        private void button13_Click(object sender, EventArgs e)
        {
            con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";
            
            cmd = new SQLiteCommand();
            cmd.Connection = con;

            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме

            con.Open(); // открываем соединение с БД            
            cmd.CommandText = "SELECT д.Наименование, count(*) AS Пропуски FROM Посещаемость п INNER JOIN Расписание р ON(п.Номер_урока = р.Номер_урока AND п.Дата = р.Дата) INNER JOIN Дисциплина д ON(р.Код_дисциплины = д.Код_дисциплины) GROUP BY д.Наименование ORDER BY count(*) DESC, д.Наименование ASC";
            dt.Load(cmd.ExecuteReader()); // выполняем SQL-запрос
            con.Close(); // закрываем соединение с БД

            dataGridView1.Columns[0].HeaderText = "Назва дисципліни";
            dataGridView1.Columns[1].HeaderText = "Пропуски";
        }
        //Пропуски студентов
        private void button14_Click(object sender, EventArgs e)
        {
            con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";

            cmd = new SQLiteCommand();
            cmd.Connection = con;

            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме

            con.Open(); // открываем соединение с БД
            cmd.CommandText = "SELECT с.Фамилия, с.Имя, count(*) AS Пропуски FROM Посещаемость п INNER JOIN Студент с ON(п.Шифр_студента = с.Шифр_студента) GROUP BY с.Фамилия, с.Имя";
            dt.Load(cmd.ExecuteReader()); // выполняем SQL-запрос
            con.Close(); // закрываем соединение с БД

            dataGridView1.Columns[0].HeaderText = "Прізвище";
            dataGridView1.Columns[1].HeaderText = "Ім'я";
            dataGridView1.Columns[2].HeaderText = "К-ть пропусків";

        }
        //Отчет обу успеваемости
        private void button15_Click(object sender, EventArgs e)
        {
            con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";

            cmd = new SQLiteCommand();
            cmd.Connection = con;

            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме

            con.Open(); // открываем соединение с БД
            cmd.CommandText = "SELECT с.Фамилия, с.Имя, у.Аттестация, д.Наименование FROM Студент с INNER JOIN Успеваемость у ON(с.Шифр_студента = у.Шифр_студента) INNER JOIN Дисциплина д ON(у.Код_дисциплины = д.Код_дисциплины) GROUP BY с.Фамилия, с.Имя,у.Аттестация, д.Наименование";
            dt.Load(cmd.ExecuteReader()); // выполняем SQL-запрос
            con.Close(); // закрываем соединение с БД

            dataGridView1.Columns[0].HeaderText = "Прізвище";
            dataGridView1.Columns[1].HeaderText = "Ім'я";
            dataGridView1.Columns[2].HeaderText = "Атестація";
            dataGridView1.Columns[3].HeaderText = "Назва дисципліни";
        }
        //Список неуспевающих
        private void button16_Click(object sender, EventArgs e)
        {
            con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";

            cmd = new SQLiteCommand();
            cmd.Connection = con;

            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме

            con.Open(); // открываем соединение с БД
            cmd.CommandText = "SELECT  с.Фамилия, с.Имя, д.Наименование AS Предмет, у.Аттестация AS Отметка FROM Успеваемость у INNER JOIN Студент с ON(у.Шифр_студента = с.Шифр_студента) INNER JOIN Дисциплина д ON(у.Код_дисциплины = д.Код_дисциплины) WHERE у.Аттестация = '2' GROUP BY у.Аттестация, с.Шифр_студента, с.Фамилия, с.Имя, д.Наименование";
            dt.Load(cmd.ExecuteReader()); // выполняем SQL-запрос
            con.Close(); // закрываем соединение с БД

            dataGridView1.Columns[0].HeaderText = "Прізвище";
            dataGridView1.Columns[1].HeaderText = "Ім'я";
            dataGridView1.Columns[2].HeaderText = "Дисципліна";
            dataGridView1.Columns[3].HeaderText = "Оцінка";
        }

        private void button17_Click(object sender, EventArgs e)
        {
            con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";

            cmd = new SQLiteCommand();
            cmd.Connection = con;

            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме

            con.Open(); // открываем соединение с БД
            cmd.CommandText = "SELECT д.Наименование AS Предмет, д.Преподаватель, з.Вид_задания, з.Дата_сдачи FROM Дисциплина д INNER JOIN Задания з ON(д.Код_дисциплины = з.Код_дисциплины) ORDER BY з.Дата_сдачи ASC";
            dt.Load(cmd.ExecuteReader()); // выполняем SQL-запрос
            con.Close(); // закрываем соединение с БД

            dataGridView1.Columns[0].HeaderText = "Дисципліна";
            dataGridView1.Columns[1].HeaderText = "Викладач";
            dataGridView1.Columns[2].HeaderText = "Вид завдання";
            dataGridView1.Columns[3].HeaderText = "Дата здачі";
        }

        private void button18_Click(object sender, EventArgs e)
        {
            con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";

            cmd = new SQLiteCommand();
            cmd.Connection = con;

            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме

            con.Open(); // открываем соединение с БД
            cmd.CommandText = "SELECT с.Фамилия, д.Наименование AS Предмет, з.Вид_задания, т.Тема, з.Дата_сдачи AS Сдать_к , т.Отметка_о_сдаче  FROM Дисциплина д  INNER JOIN Задания з ON(д.Код_дисциплины = з.Код_дисциплины) INNER JOIN Темы т ON(з.Код_задания = т.Код_задания) INNER JOIN Студент с ON(т.Шифр_студента = с.Шифр_студента) ORDER BY с.Фамилия ASC";
            dt.Load(cmd.ExecuteReader()); // выполняем SQL-запрос
            con.Close(); // закрываем соединение с БД

            dataGridView1.Columns[0].HeaderText = "Прізвище";
            dataGridView1.Columns[1].HeaderText = "Дисципліна";
            dataGridView1.Columns[2].HeaderText = "Вид завдання";
            dataGridView1.Columns[3].HeaderText = "Тема завдання";
            dataGridView1.Columns[4].HeaderText = "Здати до:";
            dataGridView1.Columns[5].HeaderText = "Відмітка про здачу";
        }

        private void файлToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // выбор файла базы данных
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Clear();
                dbName = openFileDialog1.FileName;

                textBox1.Text = dbName;
                MessageBox.Show("Файл базы данных выбран.");

            }
        }

        private void обновитьФайлБДToolStripMenuItem_Click(object sender, EventArgs e)
        {
            adapter.Update(dt); //обновляем данные в базе данных
            
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button10_Click(sender, e);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dt = new System.Data.DataTable(); // Создаем объект DataTable
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме
            con = new SQLiteConnection(); //создание нового подключения
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";
            adapter = new SQLiteDataAdapter("Select * from Студент ORDER BY Фамилия ASC", con);

            //adapter = new SQLiteDataAdapter("Select * from Студент ORDER BY Фамилия ASC", con);
            SQLiteCommandBuilder builder = new SQLiteCommandBuilder(adapter);
            try
            {
                adapter.UpdateCommand = builder.GetUpdateCommand();
                adapter.InsertCommand = builder.GetInsertCommand();
                adapter.DeleteCommand = builder.GetDeleteCommand();
                adapter.Fill(dt);
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show(ex.Message);
            }
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Прізвище";
            dataGridView1.Columns[2].HeaderText = "Ім'я";
            dataGridView1.Columns[3].HeaderText = "По батькові";
            dataGridView1.Columns[4].HeaderText = "Дата народження";
            dataGridView1.Columns[5].HeaderText = "Телефон";
            dataGridView1.Columns[6].HeaderText = "Адреса проживання";
            dataGridView1.Columns[7].HeaderText = "Відомості про батьків";
            dataGridView1.Columns[8].HeaderText = "Додаткова інформація";
            dataGridView1.Columns[9].HeaderText = "Чи є старостою";
            dataGridView1.Columns[10].HeaderText = "Форма навчання";
            dataGridView1.Columns[11].HeaderText = "Стать";
            dataGridView1.Columns[12].HeaderText = "Група";


        }

        private void button19_Click(object sender, EventArgs e)
        {



            /*con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";

            cmd = new SQLiteCommand();
            cmd.Connection = con;

            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме

            con.Open(); // открываем соединение с БД
            cmd.CommandText = "SELECT с.Фамилия, д.Наименование AS Предмет, н.Пропуск, н.Дата,  н.Номер_урока,з.Вид_задания, т.Тема, з.Дата_сдачи AS Сдать_к , т.Отметка_о_сдаче, у.Аттестация  FROM Дисциплина д, Посещаемость н INNER JOIN Успеваемость у ON(с.Шифр_студента = у.Шифр_студента)  INNER JOIN Задания з ON(д.Код_дисциплины = з.Код_дисциплины) INNER JOIN Темы т ON(з.Код_задания = т.Код_задания) INNER JOIN Студент с ON(т.Шифр_студента = с.Шифр_студента) WHERE с.Фамилия='" + textBox2.Text.Trim()+"' ORDER BY с.Фамилия ASC";
            dt.Load(cmd.ExecuteReader()); // выполняем SQL-запрос
            con.Close(); // закрываем соединение с БД
            dataGridView1.Columns[0].HeaderText = "Прізвище";
            dataGridView1.Columns[1].HeaderText = "Дисципліна";
            dataGridView1.Columns[2].HeaderText = "Пропуски";
            dataGridView1.Columns[3].HeaderText = "Дата";
            dataGridView1.Columns[4].HeaderText = "Номер лекції";
            dataGridView1.Columns[5].HeaderText = "Вид завдання";
            dataGridView1.Columns[6].HeaderText = "Тема завдання";
            dataGridView1.Columns[7].HeaderText = "Здати до";
            dataGridView1.Columns[8].HeaderText = "Відмітка про здачу";
            dataGridView1.Columns[9].HeaderText = "Атестація";
            */

            con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";

            cmd = new SQLiteCommand();
            cmd.Connection = con;

            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме

            con.Open(); // открываем соединение с БД


            cmd.CommandText = "Select * from Студент WHERE TRIM(Фамилия)  LIKE '%" + textBox2.Text.Trim() + "%' ORDER BY Фамилия ASC";

            dt.Load(cmd.ExecuteReader()); // выполняем SQL-запрос

            con.Close(); // закрываем соединение с БД

            dataGridView1.Columns[1].HeaderText = "Прізвище";
            dataGridView1.Columns[2].HeaderText = "Ім'я";
            dataGridView1.Columns[3].HeaderText = "По батькові";
            dataGridView1.Columns[4].HeaderText = "Дата народження";
            dataGridView1.Columns[5].HeaderText = "Телефон";
            dataGridView1.Columns[6].HeaderText = "Адреса проживання";
            dataGridView1.Columns[7].HeaderText = "Відомості про батьків";
            dataGridView1.Columns[8].HeaderText = "Додаткова інформація";
            dataGridView1.Columns[9].HeaderText = "Чи є старостою";
            dataGridView1.Columns[10].HeaderText = "Форма навчання";
            dataGridView1.Columns[11].HeaderText = "Стать";
            dataGridView1.Columns[12].HeaderText = "Група";



        }

        private void textBox2_Enter(object sender, EventArgs e)
        {
            textBox2.Text = "";
        }

        private void button20_Click(object sender, EventArgs e)
        {
            
        }

        private void button21_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "export.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {

                Export_Data_To_Word(dataGridView1, sfd.FileName);
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";

            cmd = new SQLiteCommand();
            cmd.Connection = con;

            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме

            con.Open(); // открываем соединение с БД


            cmd.CommandText = "Select * from Студент WHERE Староста = 1 ORDER BY Фамилия ASC";

            dt.Load(cmd.ExecuteReader()); // выполняем SQL-запрос

            con.Close(); // закрываем соединение с БД

            dataGridView1.Columns[1].HeaderText = "Прізвище";
            dataGridView1.Columns[2].HeaderText = "Ім'я";
            dataGridView1.Columns[3].HeaderText = "По батькові";
            dataGridView1.Columns[4].HeaderText = "Дата народження";
            dataGridView1.Columns[5].HeaderText = "Телефон";
            dataGridView1.Columns[6].HeaderText = "Адреса проживання";
            dataGridView1.Columns[7].HeaderText = "Відомості про батьків";
            dataGridView1.Columns[8].HeaderText = "Додаткова інформація";
            dataGridView1.Columns[9].HeaderText = "Чи є старостою";
            dataGridView1.Columns[10].HeaderText = "Форма навчання";
            dataGridView1.Columns[11].HeaderText = "Стать";
            dataGridView1.Columns[12].HeaderText = "Група";
        }

        private void button23_Click(object sender, EventArgs e)
        {
            con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";

            cmd = new SQLiteCommand();
            cmd.Connection = con;

            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме

            con.Open(); // открываем соединение с БД


            cmd.CommandText = "Select * from Студент WHERE Форма = 'бюджет' ORDER BY Фамилия ASC";

            dt.Load(cmd.ExecuteReader()); // выполняем SQL-запрос

            con.Close(); // закрываем соединение с БД

            dataGridView1.Columns[1].HeaderText = "Прізвище";
            dataGridView1.Columns[2].HeaderText = "Ім'я";
            dataGridView1.Columns[3].HeaderText = "По батькові";
            dataGridView1.Columns[4].HeaderText = "Дата народження";
            dataGridView1.Columns[5].HeaderText = "Телефон";
            dataGridView1.Columns[6].HeaderText = "Адреса проживання";
            dataGridView1.Columns[7].HeaderText = "Відомості про батьків";
            dataGridView1.Columns[8].HeaderText = "Додаткова інформація";
            dataGridView1.Columns[9].HeaderText = "Чи є старостою";
            dataGridView1.Columns[10].HeaderText = "Форма навчання";
            dataGridView1.Columns[11].HeaderText = "Стать";
            dataGridView1.Columns[12].HeaderText = "Група";
        }

        private void button24_Click(object sender, EventArgs e)
        {
            con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";

            cmd = new SQLiteCommand();
            cmd.Connection = con;

            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме

            con.Open(); // открываем соединение с БД


            cmd.CommandText = "Select * from Студент WHERE Форма = 'контракт' ORDER BY Фамилия ASC";

            dt.Load(cmd.ExecuteReader()); // выполняем SQL-запрос

            con.Close(); // закрываем соединение с БД

            dataGridView1.Columns[1].HeaderText = "Прізвище";
            dataGridView1.Columns[2].HeaderText = "Ім'я";
            dataGridView1.Columns[3].HeaderText = "По батькові";
            dataGridView1.Columns[4].HeaderText = "Дата народження";
            dataGridView1.Columns[5].HeaderText = "Телефон";
            dataGridView1.Columns[6].HeaderText = "Адреса проживання";
            dataGridView1.Columns[7].HeaderText = "Відомості про батьків";
            dataGridView1.Columns[8].HeaderText = "Додаткова інформація";
            dataGridView1.Columns[9].HeaderText = "Чи є старостою";
            dataGridView1.Columns[10].HeaderText = "Форма навчання";
            dataGridView1.Columns[11].HeaderText = "Стать";
            dataGridView1.Columns[12].HeaderText = "Група";
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void button25_Click(object sender, EventArgs e)
        {
            button15.PerformClick();
            chart1.Series[0].Points.Clear();
            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -90; 
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                string x = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value)+", "+Convert.ToString(dataGridView1.Rows[i].Cells[3].Value);
                double y = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                chart1.Series[0].Points.AddXY(x, y);
            }

        }

        private void button27_Click(object sender, EventArgs e)
        {
            button2.PerformClick();
            chart1.Series[0].Points.Clear();
            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = 0;
            int k1 = 0, k2 = 0;

            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Cells[10].Value.ToString() == "бюджет")
                    k1++;
                if (dataGridView1.Rows[i].Cells[10].Value.ToString() == "контракт")
                    k2++;
            }

            chart1.Series[0].Points.AddXY("Кількість на бюджеті ("+k1.ToString()+")", k1);
            chart1.Series[0].Points.AddXY("Кількість на контракті (" + k2.ToString() + ")", k2);
        }

        private void button26_Click(object sender, EventArgs e)
        {
            button2.PerformClick();
            chart1.Series[0].Points.Clear();
            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = 0;
            int k1 = 0, k2 = 0;

            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Cells[11].Value.ToString() == "м")
                    k1++;
                if (dataGridView1.Rows[i].Cells[11].Value.ToString() == "ж")
                    k2++;
            }

            chart1.Series[0].Points.AddXY("Кількість чоловіків (" + k1.ToString() + ")", k1);
            chart1.Series[0].Points.AddXY("Кількість жінок (" + k2.ToString() + ")", k2);
        }

        private void button28_Click(object sender, EventArgs e)
        {
            con = new SQLiteConnection();
            con.ConnectionString = @"Data Source=" + dbName + ";New=False;Version=3";

            cmd = new SQLiteCommand();
            cmd.Connection = con;

            dt = new System.Data.DataTable();
            dataGridView1.DataSource = dt; // связываем DataTable и таблицу на форме

            con.Open(); // открываем соединение с БД


            cmd.CommandText = "Select * from Студент WHERE Група = '" + textBox3.Text.Trim()+"' ORDER BY Фамилия ASC";

            dt.Load(cmd.ExecuteReader()); // выполняем SQL-запрос

            con.Close(); // закрываем соединение с БД

            dataGridView1.Columns[1].HeaderText = "Прізвище";
            dataGridView1.Columns[2].HeaderText = "Ім'я";
            dataGridView1.Columns[3].HeaderText = "По батькові";
            dataGridView1.Columns[4].HeaderText = "Дата народження";
            dataGridView1.Columns[5].HeaderText = "Телефон";
            dataGridView1.Columns[6].HeaderText = "Адреса проживання";
            dataGridView1.Columns[7].HeaderText = "Відомості про батьків";
            dataGridView1.Columns[8].HeaderText = "Додаткова інформація";
            dataGridView1.Columns[9].HeaderText = "Чи є старостою";
            dataGridView1.Columns[10].HeaderText = "Форма навчання";
            dataGridView1.Columns[11].HeaderText = "Стать";
            dataGridView1.Columns[12].HeaderText = "Група";
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            button21_Click(sender, e);
        }
    }
}
