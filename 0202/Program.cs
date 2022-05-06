using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using static System.Console;
using System.Diagnostics;
using System.IO;


namespace _0202
{
    class Program
    {


        static void Main(string[] args)
        {
            
            TextWriterTraceListener log = new TextWriterTraceListener(System.IO.File.CreateText(@"data\logi.txt"));
            Debug.Listeners.Add(log);
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
                        Trace.WriteLine($"{DateTime.Now}: Пользователь посетил справочник");
                        
                        exit_spr = false;
                        while(exit_spr == false)
                        {

                            Clear();
                            WriteLine("1 - Новая запись \n2 - Просмотр всех записей\n3 - Поиск \n4 - Удаление записи\n0 - Назад");
                            switch (ReadLine())
                            {
                                case "1":
                                    Trace.WriteLine($"{DateTime.Now}: Пользователь открыл форму создания новой записи");
                                    
                                    Clear();
                                    string ValueCell = "1";
                                    int i = 0;
                                    while (ValueCell != "")
                                    {
                                        i++;
                                        ValueCell = sheet.Cells[i, 1].Text;
                                    }
                                    for (int col = 1; col < titleName.Length + 1; col++)
                                    {
                                        
                                        Write(titleName[col - 1]); sheet.Cells[i, col] = String.Format(ReadLine());
                                    }
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    Trace.WriteLine("Пользователь создал новую запись");
                                    data_applicants.Application.ActiveWorkbook.SaveAs(path, Type.Missing,
                                  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                                  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                    WriteLine("Запись успешно создана. Нажмите любую клавишу чтобы продолжить");
                                    Console.ResetColor();
                                    ReadKey();
                                    break;
                                case "2":
                                    Trace.WriteLine($"{DateTime.Now}: Пользователь открыл форму просмотра записи");
                                    
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
                                                    WriteLine("Введены некоректные данные. Нажмите ENTER чтобы продолжить");
                                                    Console.ResetColor();
                                                    ReadLine();
                                                    break;
                                            }
                                        }

                                    }
                                    break;
                                case "3":
                                    Trace.WriteLine($"{DateTime.Now}: Пользователь открыл форму поиска записей");
                                    
                                    Clear();
                                    ValueCell = "1";
                                    i = 2;
                                    int check_search = 0;
                                    int count_search = 0;
                                    
                                    for (int col = 0; col < titleName.Length; col++)
                                    {
                                        Write(titleName[col]);
                                        search[col] = ReadLine();
                                    }
                                    for (int s = 0; s < search.Length; s++)
                                    {
                                        if (search[s] != "") count_search++;
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
                                                    check_search++;
                                                }
                                            }
                                            if (check_search == count_search)
                                            {
                                                for (int col = 0; col < titleName.Length; col++)
                                                {
                                                    WriteLine(titleName[col] + sheet.Cells[i,col+1].Text);
                                                }
                                                check_search = 0;
                                                WriteLine("_____________________________________________________________________");
                                            }

                                        }
                                        check_search = 0;
                                        i++;
                                    }

                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    WriteLine("Нажмите ENTER чтобы продолжить");
                                    Console.ResetColor();
                                    ReadLine();
                                    break;
                                case "0":
                                    exit_spr = true;
                                    break;
                                case "4":
                                    Trace.WriteLine($"{DateTime.Now}: Пользователь открыл форму удаления записи");

                                    Clear();
                                    ValueCell = "1";
                                    i = 2;
                                    check_search = 0;
                                    count_search = 0;
                                    int i_del = 0;
                                    for (int col = 0; col < titleName.Length; col++)
                                    {
                                        Write(titleName[col]);
                                        search[col] = ReadLine();
                                    }
                                    for (int s = 0; s < search.Length; s++)
                                    {
                                        if (search[s] != "") count_search++;
                                    }
                                    WriteLine("_____________________________________________________________________");
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    WriteLine("Запись:");
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
                                                    check_search++;
                                                }
                                            }
                                            if (check_search == count_search)
                                            {
                                                for (int col = 0; col < titleName.Length; col++)
                                                {
                                                    WriteLine(titleName[col] + sheet.Cells[i, col + 1].Text);
                                                }
                                                check_search = 0;
                                                WriteLine("_____________________________________________________________________");
                                                Console.ForegroundColor = ConsoleColor.Yellow;
                                                WriteLine("Удалить эту запись?         [y/n]");
                                                Console.ResetColor();
                                                check = false;
                                                while (check == false)
                                                {
                                                    switch (ReadLine())
                                                    {
                                                        case "y":
                                                            for (int col = 0; col < titleName.Length; col++)
                                                            {
                                                                sheet.Cells[i, col + 1] = "";
                                                            }
                                                            i_del = i;
                                                            check = true;
                                                            break;
                                                        case "n":
                                                            check = true;
                                                            break;
                                                        default:
                                                            Console.ForegroundColor = ConsoleColor.Yellow;
                                                            WriteLine("Введены некоректные данные. Нажмите ENTER чтобы продолжить");
                                                            Console.ResetColor();
                                                            ReadLine();
                                                            break;
                                                    }
                                                    i++;
                                                }
                                                
                                            }

                                        }
                                        check_search = 0;
                                        i++;
                                    }
                                    Excel.Range range_moved = ObjWorkSheet.get_Range("A"+(i_del+1), "O"+i);
                                    Excel.Range range_move = ObjWorkSheet.get_Range("A" + i_del, "O" + (i-1));
                                    range_moved.Cut(range_move);
                                    data_applicants.Application.ActiveWorkbook.SaveAs(path, Type.Missing,
                                  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                                  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    WriteLine("Запись удалена. Нажмите ENTER чтобы продолжить");
                                    Console.ResetColor();
                                    ReadLine();
                                    break;
                                default:
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    WriteLine("Введены некоректные данные. Нажмите ENTER чтобы продолжить");
                                    Console.ResetColor();
                                    ReadLine();
                                    break;
                            }
                        }

                        break;
                    case "2":
                        Trace.WriteLine($"{DateTime.Now}: Пользователь открыл форму просмотра отчетов");
                        
                        bool back = false;
                        while (back == false)
                        {
                            Clear();
                            WriteLine("Выберите отчет: \n\nСтатистика по специальностям - 1\nСпатистика по изучаемому ин.языку - 2");
                            WriteLine("\n Назад - 0");
                            switch (ReadLine())
                            {
                                case "1":
                                    Trace.WriteLine($"{DateTime.Now}: Пользователь открыл форму отчета по специальностям");
                                    int i = 1;
                                    int col = 2;
                                    string ValueCell = "1";
                                    Clear();
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    WriteLine("Отчет \"Статистика по специальнотсям\"");
                                    Console.ResetColor();
                                    while(sheet.Cells[col, 8].text != "")
                                    {
                                        col++;
                                        i++;
                                    }
                                    Excel.Range rangeKey = ObjWorkSheet.get_Range("H2", "H" + i);
                                    ObjWorkSheet.Sort.SortFields.Clear();
                                    ObjWorkSheet.Sort.SortFields.Add(rangeKey);
                                    ObjWorkSheet.Sort.SetRange(ObjWorkSheet.Range["H2","H" + i]);
                                    ObjWorkSheet.Sort.Orientation = Excel.XlSortOrientation.xlSortColumns;
                                    ObjWorkSheet.Sort.SortMethod = Excel.XlSortMethod.xlPinYin;
                                    ObjWorkSheet.Sort.Apply();

                                    col = 2;
                                    i = 0;
                                    ValueCell = sheet.Cells[col, 8].text;
                                    while (sheet.Cells[col, 8].text != "")
                                    {
                                        if(sheet.Cells[col, 8].text == ValueCell)
                                        {
                                            i++;
                                            if (sheet.Cells[col + 1, 8].text == "") WriteLine(ValueCell + " - " + i);
                                        }
                                        else
                                        {
                                            WriteLine(ValueCell + " - " + i);
                                            ValueCell = sheet.Cells[col, 8].text;
                                            i = 1;
                                            if(sheet.Cells[col+1, 8].text == "") WriteLine(ValueCell + " - " + i);
                                        }
                                        col++; 
                                    }
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    WriteLine("Нажмите ENTER чтобы продолжить");
                                    Console.ResetColor();
                                    ReadLine();
                                    break;
                                case "2":
                                    Trace.WriteLine($"{DateTime.Now}: Пользователь открыл форму отчета по изучаемому языку");
                                    i = 1;
                                    col = 2;
                                    ValueCell = "1";
                                    Clear();
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    WriteLine("Отчет \"Статистика по изучаемому языку\"");
                                    Console.ResetColor();
                                    while (sheet.Cells[col, 8].text != "")
                                    {
                                        col++;
                                        i++;
                                    }
                                    rangeKey = ObjWorkSheet.get_Range("N2", "N" + i);
                                    ObjWorkSheet.Sort.SortFields.Clear();
                                    ObjWorkSheet.Sort.SortFields.Add(rangeKey);
                                    ObjWorkSheet.Sort.SetRange(ObjWorkSheet.Range["N2", "N" + i]);
                                    ObjWorkSheet.Sort.Orientation = Excel.XlSortOrientation.xlSortColumns;
                                    ObjWorkSheet.Sort.SortMethod = Excel.XlSortMethod.xlPinYin;
                                    ObjWorkSheet.Sort.Apply();

                                    col = 2;
                                    i = 0;
                                    ValueCell = sheet.Cells[col, 14].text;
                                    while (sheet.Cells[col, 14].text != "")
                                    {
                                        if (sheet.Cells[col, 14].text == ValueCell)
                                        {
                                            i++;
                                            if (sheet.Cells[col + 1, 8].text == "") WriteLine(ValueCell + " - " + i);
                                        }
                                        else
                                        {
                                            WriteLine(ValueCell + " - " + i);
                                            ValueCell = sheet.Cells[col, 14].text;
                                            i = 1;
                                            if (sheet.Cells[col + 1, 8].text == "") WriteLine(ValueCell + " - " + i);
                                        }
                                        col++;
                                    }
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    WriteLine("Нажмите ENTER чтобы продолжить");
                                    Console.ResetColor();
                                    ReadLine();
                                    break;
                                case "0":
                                    back = true;
                                    break;
                                default:
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    WriteLine("Введены некоректные данные. Нажмите ENTER чтобы продолжить");
                                    Console.ResetColor();
                                    ReadLine();
                                    break;
                            }
                        }
               
                        
                        break;
                    case "3":
                        exit = true;
                        Trace.WriteLine($"{DateTime.Now}: Пользователь вышел из программы");
                        
                        break;
                    default:
                        Debug.Assert(false);
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        WriteLine("Введены некоректные данные. Нажмите ENTER чтобы продолжить");
                        Console.ResetColor();
                        ReadLine();
                        break;
                }

            }
            Debug.Flush();
            WorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            data_applicants.Quit(); // выйти из экселя
            GC.Collect(); // убрать за собой
            Environment.Exit(0);
        }
    }
}
