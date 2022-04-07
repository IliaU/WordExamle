using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Threading;
using System.Data;
using WordDotx;

namespace ExcelExample
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
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
                TabTmp.Rows.Add(nrow);
                nrow = TabTmp.NewRow();
                nrow["A"] = "1";
                nrow["B"] = "2";
                nrow["C"] = "3";
                TabTmp.Rows.Add(nrow);
                // Добавлем эту таблицу в наш класс
                Table Tab = new Table("1|B4", TabTmp);   // передаём индекс страницы (начинается с 1) и ячейку таблицы (её самый левый верхний угол) 
                TabL.Add(Tab, true);


                // Создаём задание
                TaskExcel Tsk = new TaskExcel(Environment.CurrentDirectory.Replace(@"ExcelExample\bin\Debug", @"Шаблон.xlsx"), Environment.CurrentDirectory.Replace(@"ExcelExample\bin\Debug", @"Результат.xlsx"), TabL, true);

                // запускаем асинхронно но в несколько потоков
                //List<TaskExcel> rezL = StartASinchronePool(Tsk);

                // запускаем асинхронно но в один поток
                //TaskExcel rez = StartASinchrone(Tsk);

                // Запускаем формирование отчёта в синхронном режиме
                StartSinchrone(Tsk);


                Console.WriteLine(string.Format("Success"));


                /*
                 
                // Можно смотреть версию приложения и понимать нужно ли попросить обновиться пользователя или нет
                int[] ver = FarmExcel.VersionDll;

                // Хотим создавать статичный класс которы будет обрабатывать наши объекты
                // Метод имее перегрузку можно указать входную и выходную папку по умолчанию где берём шаблоны для обработки и куда клоадём результат
                ExcelServer SrvStatic = FarmExcel.CreateExcelServer();

                // Так можно обращаться к текущему серверу если хоть раз его инициировали то он создаётся
                SrvStatic = FarmExcel.CurrentExcelServer;

                */

                //ExcelServer.OlpenReport("Путь к файлу или если вызываем метод в екземпляре то можно передать задание");
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
        private static List<TaskExcel> StartASinchronePool(TaskExcel Tsk)
        {
            FarmExcel.PoolWorkerList.onEvTaskExcelEnd += PoolWorkerList_onEvTaskExcelEnd;

            // Запускаем пул с несколькими потоками по усолчанию с количеством как физическое кол-во CPU
            FarmExcel.PoolWorkerList.Start();

            // Добавляем задание в очередь
            RezultTaskExcel rez = FarmExcel.QueTaskExcelAdd(Tsk);

            // Небольшая пауза чтобы успел завестись парралельный поток в реальной жизни не нужно так как кгда останавливаем мы не хотим дожидаться завершения всех потоков
            Thread.Sleep(1000);
                        
            // Можно смотреть что сейчас выполняется и получать стату по задачам которые сейчас в процессе
            List<TaskExcel> TskL = FarmExcel.PoolWorkerList.GetTaskExcelList();

            // Останавливаем наш пул
            FarmExcel.PoolWorkerList.Stop();
            FarmExcel.PoolWorkerList.Join();

            return TskL;
        }

        // Вот так можно подписаться и получать события когда наши задания будут выполняться
        private static void PoolWorkerList_onEvTaskExcelEnd(object sender, EvTaskExcelEnd e)
        {
            TaskExcel Tsk = e.Tsk;
        }


        /// <summary>
        /// Запускаем формирование отчёта в синхронном режиме
        /// </summary>
        /// <param name="Tsk"></param>
        private static TaskExcel StartASinchrone(TaskExcel Tsk)
        {
            WorkerExcel wrk = new WorkerExcel();
            wrk.Start();

            // Добавляем задание в очередь
            RezultTaskExcel rez = FarmExcel.QueTaskExcelAdd(Tsk);

            // Небольшая пауза чтобы успел завестись парралельный поток в реальной жизни не нужно так как кгда останавливаем мы не хотим дожидаться завершения всех потоков
            Thread.Sleep(1000);

            // Если гдето задание потеряли то его можно получить посмотрев что сейчас выполняеки работник
            TaskExcel rrr = wrk.TaskExl;

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
        private static void StartSinchrone(TaskExcel Tsk)
        {
            // Можно создать отдельный екземпляр который сможет работать асинхронно со своими параметрами
            ExcelServer SrvStatic = new ExcelServer("dd", "rrr");

            // Запускаем формирование отчёта в синхронном режиме
            SrvStatic.StartCreateReport(Tsk);
        }
    }
}
