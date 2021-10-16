using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using WordDotx;

namespace WordExamle
{
    class Program
    {
        static void Main(string[] args)
        {
            // Хотим создавать статичный класс которы будет обрабатывать наши объекты
            // Метод имее перегрузку можно указать входную и выходную папку по умолчанию где берём шаблоны для обработки и куда клоадём результат
            WordDotxServer SrvStatic = FarmWordDotx.CreateWordDotxServer();

            // Так можно обращаться к текущему серверу если хоть раз его инициировали то он создаётся
            SrvStatic = FarmWordDotx.CurrentWordDotxServer;

            // Можно создать отдельный екземпляр который сможет работать асинхронно со своими параметрами
            WordDotxServer Srv2 = new WordDotxServer("dd", "rrr");

            // Создаём список закладок
            BookmarkList BmL = new BookmarkList();
            Bookmark Bm = new Bookmark("Z1", "НОВЫЙ ТЕКСТ");
            BmL.Add(Bm, true);

            // Создаём таблицу с которой потом будем работать
            TableList TabL = new TableList();
            // Создаём временную таблицу
            DataTable TabTmp = new DataTable();
            TabTmp.Columns.Add(new DataColumn("A", typeof(string)));
            TabTmp.Columns.Add(new DataColumn("B", typeof(string)));
            TabTmp.Columns.Add(new DataColumn("C", typeof(string)));
            DataRow nrow = TabTmp.NewRow();
            nrow["A"] = "A1";
            nrow["B"] = "B1";
            nrow["C"] = "C1";
            TabTmp.Rows.Add(nrow);
            nrow = TabTmp.NewRow();
            nrow["A"] = "A2";
            nrow["B"] = "B2";
            nrow["C"] = "C2";
            TabTmp.Rows.Add(nrow);
            nrow = TabTmp.NewRow();
            nrow["A"] = "A3";
            nrow["B"] = "B3";
            nrow["C"] = "C3";
            // Добавлем эту таблицу в наш класс
            TabTmp.Rows.Add(nrow);
            Table Tab = new Table("T", TabTmp);
            TabL.Add(Tab, true);
        }
    }
}
