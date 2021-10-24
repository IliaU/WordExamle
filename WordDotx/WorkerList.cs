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
    public class WorkerList : Lib.WorkerBase.WorkerListBase
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
        public event EventHandler<EvTaskWordStart> onEvTaskWordStart;

        /// <summary>
        /// Событие выполнения задания без ошибок
        /// </summary>
        public event EventHandler<EvTaskWordEnd> onEvTaskWordEnd;

        /// <summary>
        /// Событие исключения при получении ошибки в задании
        /// </summary>
        public event EventHandler<EvTaskWordError> onEvTaskWordError;

        /// <summary>
        /// Событие исключения которое возникло в работнике и он не может продолжать обрабатывать документы
        /// </summary>
        public event EventHandler<EvWorkerError> onEvWorkerError;

        /// <summary>
        /// Событие исключения которое возникло в пуле и он не может продолжать обрабатывать документы
        /// </summary>
        public event EventHandler<EvWorkerListError> onEvWorkerListError;

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="MaxCountThreadOfPull">Максимальное количество потоков которое нужно создавать в пуле</param>
        public WorkerList(int MaxCountThreadOfPull)
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
        public WorkerList() : this(Environment.ProcessorCount)
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
                    Thr.Name = "WorkerList";

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
        public List<TaskWord> GetTaskWordList()
        {
            try
            {
                List<TaskWord> rez = new List<TaskWord>();

                lock (Lok)
                {
                    // Пробегаем по списку работников
                    for (int i = 0; i < base.Count; i++)
                    {
                        // Вытаскиваем самого работника
                        Worker wrk = (Worker)base[i];

                        // Получаем задание которое он сейчас обрабатывает
                        TaskWord Tsk = wrk.TaskWrk;

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
                throw new ApplicationException(string.Format("{0}.GetTaskWordList   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
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
                    if (this.MaxCountThreadOfPull > 0 && FarmWordDotx.QueTaskWordCount > base.Count)
                    {
                        try
                        {
                            // Если процесс ещё не остановлен
                            if (_HashRunning)
                            {
                                // Cоздаём поток серверный и запускаем его чтобы он сразу был в работе
                                Worker wrk = new Worker();
                                wrk.onEvTaskWordStart += Wrk_onEvTaskWordStart;
                                wrk.onEvTaskWordEnd += Wrk_onEvTaskWordEnd;
                                wrk.onEvTaskWordError += Wrk_onEvTaskWordError;
                                wrk.onEvWorkerError += Wrk_onEvWorkerError;
                                wrk.Start();

                                // Добавляем процесс в пул
                                base.Add(wrk, true);
                            }
                        }
                        catch (Exception ex)
                        {
                            // Передаём ошибку подписанному пользователю на событие но процесс не завершаем
                            if (this.onEvWorkerListError != null)
                            {
                                EvWorkerListError ArgErrorL = new EvWorkerListError(this, string.Format("Ошибка при добавлениии в пул работника",ex.Message));
                                this.onEvWorkerListError.Invoke(this, ArgErrorL);
                            }
                        }
                    }

                    // Если заданий меньше чем процессов то можно убить один из потоков но сначала отправляем команду стоп свободному процессу а при следующей итерации когда он остановится можно сделать Join и удалить процесс
                    if (base.Count > 0 && FarmWordDotx.QueTaskWordCount < base.Count)
                    {
                        try
                        {
                            //  берём последний в списке компонент
                            Lib.WorkerBase wrk = (Worker)base[base.Count - 1];

                            // Останавливаем в компонент чтобы он не брал больше новых заданий но не рубим его
                            if(wrk.StatusWorker != EnStatusWorkercs.Stopping 
                                && wrk.StatusWorker != EnStatusWorkercs.Stopped
                                && wrk.StatusWorker != EnStatusWorkercs.FatalError) wrk.Stop();
                        }
                        catch (Exception ex)
                        {
                            // Передаём ошибку подписанному пользователю на событие но процесс не завершаем
                            if (this.onEvWorkerListError != null)
                            {
                                EvWorkerListError ArgErrorL = new EvWorkerListError(this, string.Format("Ошибка при остановке в пуле работника", ex.Message));
                                this.onEvWorkerListError.Invoke(this, ArgErrorL);
                            }
                        }
                    }

                    // Если есть компоненты и последний компонент в статусе остановлен значит его можно убить ждём завершения его работы и убиваем его
                    if (base.Count > 0 && FarmWordDotx.QueTaskWordCount < base.Count &&  
                        (base[base.Count - 1].StatusWorker == EnStatusWorkercs.Stopped 
                        || base[base.Count - 1].StatusWorker == EnStatusWorkercs.FatalError))
                    {
                        try
                        {

                            Lib.WorkerBase wrk = (Worker)base[base.Count - 1];
                            wrk.Join();
                            base.Remove(wrk, true);
                        }
                        catch (Exception ex)
                        {
                            // Передаём ошибку подписанному пользователю на событие но процесс не завершаем
                            if (this.onEvWorkerListError != null)
                            {
                                EvWorkerListError ArgErrorL = new EvWorkerListError(this, string.Format("Ошибка при уничтожении в пуле работника", ex.Message));
                                this.onEvWorkerListError.Invoke(this, ArgErrorL);
                            }
                        }
                    }

                    // производим лечение потоков если упал какой-то из потоков то пробуем его уничтожить чтобы создался девственно чистый поток
                    for (int i = 0; i < base.Count; i++)
                    {
                        try
                        {
                            Lib.WorkerBase wrk = base[i];

                            // Если поток упал но он не требует уничтожения потока при выключении то можно попробовать его уничтожить чтобы создался новый девственный поток
                            if (!wrk.HashRunning && wrk.StatusWorker == EnStatusWorkercs.Stopped)
                            {
                                wrk.Join();
                                base.Remove(wrk, true);
                            }
                        }
                        catch (Exception ex)
                        {
                            // Передаём ошибку подписанному пользователю на событие но процесс не завершаем
                            if (this.onEvWorkerListError != null)
                            {
                                EvWorkerListError ArgErrorL = new EvWorkerListError(this, string.Format("Ошибка при лечении в пуле работника", ex.Message));
                                this.onEvWorkerListError.Invoke(this, ArgErrorL);
                            }
                        }
                    }

                    // Если небыло команды по остановке  потока и если заданий нет то пауза
                    if (_HashRunning) Thread.Sleep(FarmWordDotx.TimeoutForWorkerSec * 1000);
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
                if (this.onEvWorkerListError != null)
                {
                    EvWorkerListError ArgErrorL = new EvWorkerListError(this, ex.Message);
                    this.onEvWorkerListError.Invoke(this, ArgErrorL);
                }

                throw new ApplicationException(string.Format("{0}.Run   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Подписка на исключния вызванные в работнике
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Wrk_onEvWorkerError(object sender, EvWorkerError e)
        {
            try
            {
                if (this.onEvWorkerError != null)
                {
                    this.onEvWorkerError.Invoke(this, e);
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.Wrk_onEvWorkerError   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Подписка связанная с получением исключения при выполнении задания
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Wrk_onEvTaskWordError(object sender, EvTaskWordError e)
        {
            try
            {
                if (this.onEvTaskWordError != null)
                {
                    this.onEvTaskWordError.Invoke(this, e);
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.Wrk_onEvTaskWordError   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Подписка на событие связанное с успешным выполнением задания
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Wrk_onEvTaskWordEnd(object sender, EvTaskWordEnd e)
        {
            try
            {
                if (this.onEvTaskWordEnd != null)
                {
                    this.onEvTaskWordEnd.Invoke(this, e);
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.Wrk_onEvTaskWordEnd   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Подписка на события связанные с запуском заданий
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Wrk_onEvTaskWordStart(object sender, EvTaskWordStart e)
        {
            try
            {
                if (this.onEvTaskWordStart != null)
                {
                    this.onEvTaskWordStart.Invoke(this, e);
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.Wrk_onEvTaskWordEnd   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
            }
        }

    }
}
