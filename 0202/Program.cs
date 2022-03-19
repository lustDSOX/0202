using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using static System.Console;

namespace _0202
{
    class Program
    {

        static void Main(string[] args)
        {
            var path = System.IO.Path.GetFullPath(@"data\data_appl.xlsx");
            Excel.Application data_applicants = new Excel.Application(); //открыть эксель
            Excel.Workbook WorkBook = data_applicants.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)data_applicants.Sheets[1]; //получить 1 лист
            data_applicants.Visible = true;
            data_applicants.DisplayAlerts = false;
            Excel.Worksheet sheet = (Excel.Worksheet)WorkBook.Sheets[1];
            string[] titleName = new string[] { "Фамилия:","Имя:","Отчество:", "Дата рождения:" , "Гражданство:", "Пол:", "Домашний адрес:", "Специальность:", "Телефон:", "Законченное образовательное учреждение:", "Год окончания:" , "Данные о родителях:" , "Доп.сведения:", "Изучаемый ин.язык:" ,"Средний балл аттестата:"};
            string[] search = new string[titleName.Length];
            bool exit = false;
            bool exit_spr = false;
            while(exit == false)
            {
                Clear();
                WriteLine("1 - Справочник \n2 - Отчет \n3 - Выход");
                switch (ReadLine())
                {
                    case "1":
                        exit_spr = false;
                        while(exit_spr == false)
                        {

                            Clear();
                            WriteLine("1 - Новая запись \n2 - Просмотр всех записей \n3 - Поиск \n0 - Назад");
                            switch (ReadLine())
                            {
                                case "1":
                                    Clear();
                                    string ValueCell = "1";
                                    int i = 0;
                                    while (ValueCell != "")
                                    {
                                        i++;
                                        ValueCell = sheet.Cells[i, 1].Text;
                                    }
                                    for (int col = 1; col <= titleName.Length + 1; col++)
                                    {
                                        Write(titleName[col - 1]); sheet.Cells[i, col] = String.Format(ReadLine());
                                    }
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    WriteLine("Запись успешно создана. Нажмите любую клавишу чтобы продолжить");
                                    Console.ResetColor();
                                    ReadKey();
                                    break;
                                case "2":
                                    Clear();
                                    i = 1;
                                    ValueCell = "1";
                                    bool nextsrt = true;
                                    bool check = false;
                                    while (nextsrt)
                                    {
                                        check = false;
                                        Clear();
                                        for (int size = 0; size < 20; size++)
                                        {
                                            i++;
                                            ValueCell = sheet.Cells[i, 1].Text;
                                            if (ValueCell != "")
                                            {
                                                for (int col = 1; col <= titleName.Length; col++)
                                                {
                                                    WriteLine(titleName[col - 1] + sheet.Cells[i, col].Text);
                                                }
                                                WriteLine("_____________________________________________________________________");
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }
                                        while (check == false)
                                        {
                                            Console.ForegroundColor = ConsoleColor.Yellow;
                                            WriteLine("1 - Следующая страница \n2 - Назад");
                                            Console.ResetColor();
                                            switch (ReadLine())
                                            {
                                                case "1":
                                                    Clear();
                                                    check = true;
                                                    nextsrt = true;
                                                    break;
                                                case "2":
                                                    Clear();
                                                    check = true;
                                                    nextsrt = false;
                                                    break;
                                                default:
                                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                                    WriteLine("Введены некоректные данные. Нажмите любую клавишу чтобы продолжить");
                                                    Console.ResetColor();
                                                    ReadKey();
                                                    break;
                                            }
                                        }

                                    }
                                    break;
                                case "3":
                                    Clear();
                                    ValueCell = "1";
                                    i = 2;
                                    for (int col = 0; col < titleName.Length; col++)
                                    {
                                        Write(titleName[col]);
                                        search[col] = ReadLine();
                                    }
                                    WriteLine("_____________________________________________________________________");
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    WriteLine("Результыты поиска:");
                                    Console.ResetColor();
                                    while (ValueCell != "")
                                    {
                                        for (int s = 0; s < search.Length; s++)
                                        {
                                            ValueCell = sheet.Cells[i, s + 1].Text;
                                            if (ValueCell == "")
                                            {
                                                break;
                                            }

                                            if (search[s] != "")
                                            {
                                                if (sheet.Cells[i, s + 1].Text == search[s])
                                                {

                                                    for (int col = 1; col <= titleName.Length; col++)
                                                    {
                                                        WriteLine(titleName[col - 1] + sheet.Cells[i, col].Text);
                                                    }
                                                    WriteLine("_____________________________________________________________________");
                                                }
                                            }


                                        }
                                        i++;

                                    }

                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    WriteLine("Нажмите любую клавишу чтобы продолжить");
                                    Console.ResetColor();
                                    ReadKey();
                                    break;

                                case "0":
                                    exit_spr = true;
                                    break;
                                default:
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    WriteLine("Введены некоректные данные. Нажмите любую клавишу чтобы продолжить");
                                    Console.ResetColor();
                                    ReadKey();
                                    break;
                            }
                        }

                        break;
                    case "2":
                        bool back = false;
                        while (back == false)
                        {
                            Clear();
                            WriteLine("Выберите отчет: \nСтатистика по специальностям - 1\n Спатистика по изучаемому ин.языку - 2");
                            WriteLine("\n Назад - 0");
                            switch (ReadLine())
                            {
                                case "1":
                                    Clear();
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    WriteLine("Отчет \"Статистика по специальнотсям\"");
                                    Console.ResetColor();
                                    int count = 0;
                                    string ValuesCell = "1";
                                    Microsoft.Office.Interop.Excel.Range rangeKey = ObjWorkSheet.get_Range("B" + (i - 1));
                                    while (ValuesCell != "")
                                    {
                                    }
                                    break;
                                case "2":
                                    break;
                                case "0":
                                    back = true;
                                    break;
                                default:
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    WriteLine("Введены некоректные данные. Нажмите любую клавишу чтобы продолжить");
                                    Console.ResetColor();
                                    ReadKey();
                                    break;
                            }
                        }
               
                        
                        break;
                    case "3":
                        exit = true;
                        break;
                    default:
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        WriteLine("Введены некоректные данные. Нажмите любую клавишу чтобы продолжить");
                        Console.ResetColor();
                        ReadKey();
                        break;
                }

            }
            WorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            data_applicants.Quit(); // выйти из экселя
            GC.Collect(); // убрать за собой
            Environment.Exit(0);

        }
    }
}
