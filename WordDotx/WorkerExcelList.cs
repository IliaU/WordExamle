using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Threading;

namespace WordDotx
{
    /// <summary>
    /// Класс который представляет из себя бул работников и может выполнять рабуту асинхронно сразу в несколько парралелей
    /// </summary>
    public class WorkerExcelList : Lib.WorkerExcelBase.WorkerExcelListBase
    {
        // Объект асинхронного процесса
        private Thread Thr;

        /// <summary>
        /// Объект для синхронизации потоков
        /// </summary>
        private object Lok = new object();

        /// <summary>
        /// Орьект который показывает работает процесс или нет
        /// </summary>
        private bool _HashRunning = false;

        /// <summary>
        /// Максимальное количество потоков которое нужно создавать в пуле
        /// </summary>
        public int MaxCountThreadOfPull;

        /// <summary>
        /// Орьект который показывает работает процесс или нет
        /// </summary>
        public bool HashRunning
        {
            get
            {
                lock (Lok)
                {
                    return _HashRunning;
                }
            }
            private set { }
        }

        /// <summary>
        /// Событие запуска задания в работу
        /// </summary>
        public event EventHandler<EvTaskExcelStart> onEvTaskExcelStart;

        /// <summary>
        /// Событие выполнения задания без ошибок
        /// </summary>
        public event EventHandler<EvTaskExcelEnd> onEvTaskExcelEnd;

        /// <summary>
        /// Событие исключения при получении ошибки в задании
        /// </summary>
        public event EventHandler<EvTaskExcelError> onEvTaskExcelError;

        /// <summary>
        /// Событие исключения которое возникло в работнике и он не может продолжать обрабатывать документы
        /// </summary>
        public event EventHandler<EvWorkerExcelError> onEvWorkerExcelError;

        /// <summary>
        /// Событие исключения которое возникло в пуле и он не может продолжать обрабатывать документы
        /// </summary>
        public event EventHandler<EvWorkerExcelListError> onEvWorkerExcelListError;

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="MaxCountThreadOfPull">Максимальное количество потоков которое нужно создавать в пуле</param>
        public WorkerExcelList(int MaxCountThreadOfPull)
        {
            try
            {
                this.MaxCountThreadOfPull = MaxCountThreadOfPull;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}   Упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
            }
        }
        //
        /// <summary>
        /// Конструктор
        /// </summary>
        public WorkerExcelList() : this(Environment.ProcessorCount)
        {
            try
            {

            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}   Упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
            }
        }


        /// <summary>
        /// Команда запуска асинхронного процесса
        /// </summary>
        public void Start()
        {
            try
            {
                lock (Lok)
                {
                    // Проверем на то что выполняет сейчас процесс какое нибудь задание или нет
                    if (Thr != null) throw new ApplicationException("Не возможно запустить процесс так как он выполняет другую задачу.");

                    //new ThreadStart(TaskThread.Run)
                    Thr = new Thread(Run);
                    Thr.IsBackground = true;
                    Thr.Name = "WorkerListExcel";

                    _HashRunning = true;
                }

                Thr.Start();
            }
            catch (Exception ex)
            {
                lock (Lok)
                {
                    _HashRunning = false;
                }
                throw new ApplicationException(string.Format("{0}.Start   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Команда остановки асинхронного потока для того чтобы можно было потом просто асинхронно проверять необъодимость уничтожения процесса
        /// </summary>
        public void Stop()
        {
            try
            {
                lock (Lok)
                {
                    _HashRunning = false;
                    for (int i = 0; i < base.Count; i++)
                    {
                        base[i].Stop();
                    }
                }
            }
            catch (Exception ex)
            {
                lock (Lok)
                {
                    _HashRunning = false;
                }
                throw new ApplicationException(string.Format("{0}.Join   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Остановка потока в синхронном режиме с физическим завершением процессов с физическим завершением процесса
        /// </summary>
        public void Join()
        {
            try
            {
                lock (Lok)
                {
                    _HashRunning = false;
                    for (int i = 0; i < base.Count; i++)
                    {
                        base[i].Stop();
                    }

                    // Проверем на то что выполняет сейчас процесс какое нибудь задание или нет
                    if (Thr != null)
                    {
                        for (int i = 0; i < base.Count; i++)
                        {
                            base[i].Join();
                        }
                        Thr.Join();
                    }

                    // Освободим переменную
                    Thr = null;
                }
            }
            catch (Exception ex)
            {
                lock (Lok)
                {
                    _HashRunning = false;
                }
                throw new ApplicationException(string.Format("{0}.Join   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Получаем список заданий которые выполняет наш пул
        /// </summary>
        /// <returns>Список заданий из нашего пула который сейчас в работе со статистикой</returns>
        public List<TaskExcel> GetTaskExcelList()
        {
            try
            {
                List<TaskExcel> rez = new List<TaskExcel>();

                lock (Lok)
                {
                    // Пробегаем по списку работников
                    for (int i = 0; i < base.Count; i++)
                    {
                        // Вытаскиваем самого работника
                        WorkerExcel wrk = (WorkerExcel)base[i];

                        // Получаем задание которое он сейчас обрабатывает
                        TaskExcel Tsk = wrk.TaskExl;

                        // Добавляем это задание в результатирующий список
                        if (Tsk != null) rez.Add(Tsk);
                    }
                }

                return rez;
            }
            catch (Exception ex)
            {
                lock (Lok)
                {
                }
                throw new ApplicationException(string.Format("{0}.GetTaskExcelList   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Асинхронный процесс который выполняет задание
        /// </summary>
        private void Run()
        {
            try
            {
                // синхронизация потока происходит при публичном методе который читает внутреннюю переменную
                while (_HashRunning)
                {

                    // Если текущее количество потоков меньше максимально возможного и если заданий больше чем текущее количество значит можно добавить поток и запустить его
                    try
                    {
                        if (this.MaxCountThreadOfPull > base.Count && FarmExcel.QueTaskExcelCount > 0)
                        {

                            // Если процесс ещё не остановлен
                            if (_HashRunning)
                            {
                                // Cоздаём поток серверный и запускаем его чтобы он сразу был в работе
                                WorkerExcel wrk = new WorkerExcel();
                                wrk.onEvTaskExcelStart += Wrk_onEvTaskExcelStart;
                                wrk.onEvTaskExcelEnd += Wrk_onEvTaskExcelEnd;
                                wrk.onEvTaskExcelError += Wrk_onEvTaskExcelError;
                                wrk.onEvWorkerExcelError += Wrk_onEvWorkerExcelError;
                                wrk.Start();

                                // Добавляем процесс в пул
                                base.Add(wrk, true);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Передаём ошибку подписанному пользователю на событие но процесс не завершаем
                        if (this.onEvWorkerExcelListError != null)
                        {
                            EvWorkerExcelListError ArgErrorL = new EvWorkerExcelListError(this, string.Format("Ошибка при добавлениии в пул работника", ex.Message));
                            this.onEvWorkerExcelListError.Invoke(this, ArgErrorL);
                        }
                    }

                    // Если заданий меньше чем процессов то можно убить один из потоков но сначала отправляем команду стоп свободному процессу а при следующей итерации когда он остановится можно сделать Join и удалить процесс
                    try
                    {
                        if (base.Count > 0 && FarmExcel.QueTaskExcelCount < base.Count)
                        {

                            //  берём последний в списке компонент
                            Lib.WorkerExcelBase wrk = (WorkerExcel)base[base.Count - 1];

                            // Останавливаем в компонент чтобы он не брал больше новых заданий но не рубим его
                            if (wrk.StatusWorker == EnStatusWorkercs.Waiting) wrk.Stop();
                        }
                    }
                    catch (Exception ex)
                    {
                        // Передаём ошибку подписанному пользователю на событие но процесс не завершаем
                        if (this.onEvWorkerExcelListError != null)
                        {
                            EvWorkerExcelListError ArgErrorL = new EvWorkerExcelListError(this, string.Format("Ошибка при остановке в пуле работника", ex.Message));
                            this.onEvWorkerExcelListError.Invoke(this, ArgErrorL);
                        }
                    }

                    // производим лечение потоков если упал какой-то из потоков то пробуем его уничтожить чтобы создался девственно чистый поток
                    try
                    {
                        for (int i = 0; i < base.Count; i++)
                        {
                            Lib.WorkerExcelBase wrk = base[i];

                            // Если поток упал но он не требует уничтожения потока при выключении то можно попробовать его уничтожить чтобы создался новый девственный поток
                            if (!wrk.HashRunning && wrk.StatusWorker == EnStatusWorkercs.Stopped)
                            {
                                wrk.Stop();

                                wrk.Join();

                                base.Remove(wrk, true);
                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        // Передаём ошибку подписанному пользователю на событие но процесс не завершаем
                        if (this.onEvWorkerExcelListError != null)
                        {
                            EvWorkerExcelListError ArgErrorL = new EvWorkerExcelListError(this, string.Format("Ошибка при лечении в пуле работника", ex.Message));
                            this.onEvWorkerExcelListError.Invoke(this, ArgErrorL);
                        }
                    }

                    // Если небыло команды по остановке  потока и если заданий нет то пауза
                    if (_HashRunning) Thread.Sleep(FarmExcel.TimeoutForWorkerSec * 1000);
                }

                // Если цыкл остановлен то первое нужно остановить все потоки которые были созданы
                for (int i = 0; i < base.Count; i++)
                {
                    base[i].Stop();
                }

                // и уничтожить их
                for (int i = 0; i < base.Count; i++)
                {
                    base[i].Join();
                }
            }
            catch (Exception ex)
            {
                lock (Lok)
                {
                    _HashRunning = false;
                }

                // Подписка на события
                if (this.onEvWorkerExcelListError != null)
                {
                    EvWorkerExcelListError ArgErrorL = new EvWorkerExcelListError(this, ex.Message);
                    this.onEvWorkerExcelListError.Invoke(this, ArgErrorL);
                }

                throw new ApplicationException(string.Format("{0}.Run   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Подписка на исключния вызванные в работнике
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Wrk_onEvWorkerExcelError(object sender, EvWorkerExcelError e)
        {
            try
            {
                if (this.onEvWorkerExcelError != null)
                {
                    this.onEvWorkerExcelError.Invoke(this, e);
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.Wrk_onEvWorkerExcelError   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Подписка связанная с получением исключения при выполнении задания
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Wrk_onEvTaskExcelError(object sender, EvTaskExcelError e)
        {
            try
            {
                if (this.onEvTaskExcelError != null)
                {
                    this.onEvTaskExcelError.Invoke(this, e);
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.Wrk_onEvTaskExcelError   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Подписка на событие связанное с успешным выполнением задания
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Wrk_onEvTaskExcelEnd(object sender, EvTaskExcelEnd e)
        {
            try
            {
                if (this.onEvTaskExcelEnd != null)
                {
                    this.onEvTaskExcelEnd.Invoke(this, e);
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.Wrk_onEvTaskExcelEnd   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Подписка на события связанные с запуском заданий
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Wrk_onEvTaskExcelStart(object sender, EvTaskExcelStart e)
        {
            try
            {
                if (this.onEvTaskExcelStart != null)
                {
                    this.onEvTaskExcelStart.Invoke(this, e);
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.Wrk_onEvTaskExcelStart   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
            }
        }
    }
}
