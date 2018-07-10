using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;



//Курс: Доступ к источникам данных с использованием ADO.NET
//
//Тема: Введение в LINQ.Использование LINQ и коллекций
//
//1.	Реализовать класс таблицы – Area.
//2.	Реализовать метод по выгрузке данных из БД.
//3.	Создать метод, который возвращает данные в виде Array.
//4.	Создать метод, который возвращает данные в виде List<Area> 
//5.	Реализовать справочник, который возвращает ID зоны/участка, и IP адрес данной зоны/участка.Так же необходимо исключить зоны/участки у которых не заполнено поле IP
//6.	Реализовать справочник, который возвращает IP адрес и касс Area, исключить все зоны/участки, у которых отсутствует IPадрес, а так же исключить все дочерние зоны/участки(ParentId!=0)
//7.	Используя коллекцию Lookup, вернуть следующие данные.В качестве ключа использовать IP адрес, в качестве значения использовать класс Area
//8.	Вернуть первую запись из последовательности, где HiddenArea = 1
//9.	Вернуть последнюю запись из таблицы Area, указав следующий фильтр – PavilionId = 1
//10.	Используя квантификаторы, вывесит на экран значения следующих фильтров:
//a.Есть ли в таблице зоны/участки для PavilionId = 1 и IP = 10.53.34.85, 10.53.34.77, 10.53.34.53
//b.Содержатся ли данные в таблице Area с наименованием зон/участков - PT disassembly, Engine testing
//11.	Вывести сумму всех работающих работников(WorkingPeople) на зонах


namespace Xml_To_Sql
{
    class Program
    {
        static MCS db = new MCS();
        static void Main(string[] args)
        {
            ExcelPackage exp = new ExcelPackage();
            ExcelWorksheet worksheet = exp.Workbook.Worksheets.Add("List1");

            int row = 2;
            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Name";
            worksheet.Column(2).Width = 50;
            worksheet.Column(3).Width = 15;
            worksheet.Cells[1, 3].Value = "IP";

            foreach (Area  area in db.Area)
            {
                worksheet.Cells[row, 1].Value = area.AreaId;
                worksheet.Cells[row, 2].Value = area.FullName;
                worksheet.Cells[row, 3].Value = area.IP;

                row++;
            }

            Dictionary<string, Area> dicIP = db.Area.Where(w => !string.IsNullOrEmpty(w.IP) && w.ParentId != 0).Select(s =>new { s.IP}).Distinct().Select(s=>new { ip=s.IP, area = db.Area.FirstOrDefault(f=>f.IP == s.IP)}).ToDictionary(d => d.ip, d => d.area);

            ExcelWorksheet worksheet2 = exp.Workbook.Worksheets.Add("List2");

            row = 2;

            foreach (var item in dicIP)
            {
                worksheet2.Cells[row, 1].Value = item.Key;
                worksheet2.Cells[row, 2].Value = item.Value.Name;
                row++;
            }

            ILookup<string, Area> lkp = db.Area.ToLookup(l => l.IP, l=>l);

            FileStream fs = File.Create("Excel.xlsx");
            
            exp.SaveAs(fs);
        }
    }
}

