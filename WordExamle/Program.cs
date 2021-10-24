using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

using System.Data;
using WordDotx;

namespace WordExamle
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {              
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
                Table Tab = new Table("Tab", TabTmp);
                TabL.Add(Tab, true);

                // Добавляем тоталов в нашу таблицу
                Tab.TtlList.Add(new Total("Total0", "Итог по Total0---"), false);
                Tab.TtlList.Add(new Total("Total1", "Итог по Total1---"), false);
                Tab.TtlList.Add(new Total("Total2", "Итог по Total2---"), false);
                Tab.TtlList.Add(new Total("Total3", "Итог по Total3---"), false);

                // Создаём задание
                TaskWord Tsk = new TaskWord(Environment.CurrentDirectory.Replace(@"WordExamle\bin\Debug", @"Шаблон.dotx"), Environment.CurrentDirectory.Replace(@"WordExamle\bin\Debug", @"Результат.doc"), BmL, TabL, true);

                // запускаем асинхронно но в несколько потоков
                List<TaskWord>  rezL = StartASinchronePool(Tsk);

                // запускаем асинхронно но в один поток
                //TaskWord rez = StartASinchrone(Tsk);

                // Запускаем формирование отчёта в синхронном режиме
                //StartSinchrone(Tsk);


                Console.WriteLine(string.Format("Success"));


                /*
                 
                // Можно смотреть версию приложения и понимать нужно ли попросить обновиться пользователя или нет
                int[] ver = FarmWordDotx.VersionDll;

                // Хотим создавать статичный класс которы будет обрабатывать наши объекты
                // Метод имее перегрузку можно указать входную и выходную папку по умолчанию где берём шаблоны для обработки и куда клоадём результат
                WordDotxServer SrvStatic = FarmWordDotx.CreateWordDotxServer();

                // Так можно обращаться к текущему серверу если хоть раз его инициировали то он создаётся
                SrvStatic = FarmWordDotx.CurrentWordDotxServer;

                 */
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("ERROR: {0}", ex.Message));
            }

            Console.ReadLine();
        }

        /// <summary>
        /// Запускаем формирование отчёта в синхронном режиме
        /// </summary>
        /// <param name="Tsk"></param>
        private static List<TaskWord> StartASinchronePool(TaskWord Tsk)
        {
            FarmWordDotx.PoolWorkerList.onEvTaskWordEnd += PoolWorkerList_onEvTaskWordEnd;

            // Запускаем пул с несколькими потоками по усолчанию с количеством как физическое кол-во CPU
            FarmWordDotx.PoolWorkerList.Start();

            // Добавляем задание в очередь
            RezultTask rez = FarmWordDotx.QueTaskWordAdd(Tsk);

            // Небольшая пауза чтобы успел завестись парралельный поток в реальной жизни не нужно так как кгда останавливаем мы не хотим дожидаться завершения всех потоков
            Thread.Sleep(1000);


/*

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
            Table Tab = new Table("Tab", TabTmp);
            TabL.Add(Tab, true);

            // Добавляем тоталов в нашу таблицу
            Tab.TtlList.Add(new Total("Total0", "Итог по Total0---"), false);
            Tab.TtlList.Add(new Total("Total1", "Итог по Total1---"), false);
            Tab.TtlList.Add(new Total("Total2", "Итог по Total2---"), false);
            Tab.TtlList.Add(new Total("Total3", "Итог по Total3---"), false);




            // Создаём задание
            TaskWord Tsk1 = new TaskWord(Environment.CurrentDirectory.Replace(@"WordExamle\bin\Debug", @"Шаблон.dotx"), Environment.CurrentDirectory.Replace(@"WordExamle\bin\Debug", @"Результат.doc"), BmL, TabL, true);

            // Добавляем задание в очередь
            RezultTask rez1 = FarmWordDotx.QueTaskWordAdd(Tsk1);


            Thread.Sleep(300000);
            */
            // Можно смотреть что сейчас выполняется и получать стату по задачам которые сейчас в процессе
            List<TaskWord> TskL = FarmWordDotx.PoolWorkerList.GetTaskWordList();

            // Останавливаем наш пул
            FarmWordDotx.PoolWorkerList.Stop();
            FarmWordDotx.PoolWorkerList.Join();

            return TskL;
        }
        // Вот так можно подписаться и получать события когда наши задания будут выполняться
        private static void PoolWorkerList_onEvTaskWordEnd(object sender, EvTaskWordEnd e)
        {
            TaskWord Tsk = e.Tsk;
        }


        /// <summary>
        /// Запускаем формирование отчёта в синхронном режиме
        /// </summary>
        /// <param name="Tsk"></param>
        private static TaskWord StartASinchrone(TaskWord Tsk)
        {
            Worker wrk = new Worker();
            wrk.Start();

            // Добавляем задание в очередь
            RezultTask rez = FarmWordDotx.QueTaskWordAdd(Tsk);

            // Небольшая пауза чтобы успел завестись парралельный поток в реальной жизни не нужно так как кгда останавливаем мы не хотим дожидаться завершения всех потоков
            Thread.Sleep(1000);

            // Если гдето задание потеряли то его можно получить посмотрев что сейчас выполняеки работник
            TaskWord rrr = wrk.TaskWrk;

            // А уже в любом задании можно получить результат для того ытобы посмотреть на какой стадии работа
            rez = Tsk.RezTsk;
                
            // Команду 
            wrk.Stop();
            wrk.Join();

            return rrr;
        }


        /// <summary>
        /// Запускаем формирование отчёта в синхронном режиме
        /// </summary>
        /// <param name="Tsk"></param>
        private static void StartSinchrone(TaskWord Tsk)
        {
            // Можно создать отдельный екземпляр который сможет работать асинхронно со своими параметрами
            WordDotxServer SrvStatic = new WordDotxServer("dd", "rrr");

            // Запускаем формирование отчёта в синхронном режиме
            SrvStatic.StartCreateReport(Tsk);
        }
    }
}
