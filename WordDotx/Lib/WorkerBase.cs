using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

using System.Collections;

namespace WordDotx.Lib
{
    /// <summary>
    /// Класс который организует асинхронное выполнение заданий к WordDotxServer
    /// </summary>
    public abstract class WorkerBase : TaskWordBase.FarmWordDotxBase.WorkerBaseInclude
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
        /// Сервер который будет обрабатывать запросы
        /// </summary>
        protected WordDotxServer WrdDotxSrv { get; private set; }

        /// <summary>
        /// Ссылка на задание по которому идёт расчёт результата
        /// </summary>
        protected TaskWord TaskWrk { get; private set; }

        /// <summary>
        /// Событие исключения которое возникло в работнике и он не может продолжать обрабатывать документы
        /// </summary>
        protected event EventHandler<EvWorkerBaseError> onEvWorkerBaseError;

        /// <summary>
        /// Индекс элемента в списке
        /// </summary>
        public int Index { get; private set; }

        /// <summary>
        /// Статус в котором может находиться работник
        /// </summary>
        public EnStatusWorkercs StatusWorker { get; private set; }

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
        /// Конструктор
        /// </summary>
        public WorkerBase()
        {
            try
            {
                this.StatusWorker = EnStatusWorkercs.Created;
                this.Index = -1;
                this.WrdDotxSrv = new WordDotxServer(FarmWordDotx.DefaultPathSource, FarmWordDotx.DefaultPathTarget, FarmWordDotx.DefReplaseFileTarget);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}   Упали с ошибкой в конструкторе: ({1})", this.GetType().Name, ex.Message));
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

                    this.StatusWorker = EnStatusWorkercs.Waiting;

                    //new ThreadStart(TaskThread.Run)
                    Thr = new Thread(Run);
                    Thr.IsBackground = true;
                    Thr.Name = string.Format("TaskWordBase {0}", this.Index);

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
                    this.StatusWorker = EnStatusWorkercs.Stopping;
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

                    // Проверем на то что выполняет сейчас процесс какое нибудь задание или нет
                    if (Thr != null) Thr.Join();

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
        /// Асинхронный процесс который выполняет задание
        /// </summary>
        private void Run()
        {
            try
            {
                // синхронизация потока происходит при публичном методе который читает внутреннюю переменную
                while (_HashRunning)
                {
                    // Получаем задание из общей очереди
                    this.TaskWrk = base.QueTaskWordGet();

                    // Если есть задание
                    if (this.TaskWrk != null)
                    {
                        this.StatusWorker = EnStatusWorkercs.Running;

                        // Проверяем подписку если пользователь подписан на это событие передаём ему управление на время
                        if (this.onEvTaskWordStart != null)
                        {
                            EvTaskWordStart ArgStart = new EvTaskWordStart(this.TaskWrk, this.WrdDotxSrv);
                            this.onEvTaskWordStart.Invoke(this, ArgStart);
                        }

                        // Получаем объект с результатом для того чтобы можно было его править и передавать результаты пользователю
                        // Например пользователь либо сам проверяет события периодически в нашем классе либо сможет подписаться на события в буле и получать результат непосредственно по событию
                        // В идеале уже со ссылкой на свой результат и  со ссылкой на задание
                        TaskWordBase.RezultTaskBase RezTsk = base.GetRezult(this.TaskWrk);

                        // оборачиваем наш запуск для того чтобы в случае ошибки поток не падал а вызывал потом событие и сообщил пользователю
                        try
                        {
                            // Меняем статус на запущенный чтобы сервер не руганулся на то что надо запукать задание асинхронно
                            base.SetStatusTaskWord(this.TaskWrk, EnStatusTask.Running);

                            // Запускаем обработку задания в нашем процессе
                            WrdDotxSrv.StartCreateReport(this.TaskWrk);

                            // Проверяем подписку если пользователь подписан на это событие передаём ему управление на время
                            if (this.onEvTaskWordEnd != null)
                            {
                                EvTaskWordEnd ArgEnd = new EvTaskWordEnd(this.TaskWrk, this.WrdDotxSrv);
                                this.onEvTaskWordEnd.Invoke(this, ArgEnd);
                            }
                        }
                        catch (Exception ex)
                        {
                            // Ну и на последок можно записать ошибку в результат
                            base.SetStatusTaskWord(this.TaskWrk, EnStatusTask.ERROR);

                            // Проверяем подписку если пользователь подписан на это событие передаём ему управление на время
                            if (this.onEvTaskWordError != null)
                            {
                                EvTaskWordError ArgError = new EvTaskWordError(this.TaskWrk, this.WrdDotxSrv, ex.Message);
                                this.onEvTaskWordError.Invoke(this, ArgError);
                            }
                        }

                        this.StatusWorker = EnStatusWorkercs.Waiting;
                    }

                    // Если небыло команды по остановке  потока и если заданий нет то пауза
                    if (_HashRunning && this.TaskWrk == null) Thread.Sleep(FarmWordDotx.TimeoutForWorkerSec * 1000);
                }
                this.StatusWorker = EnStatusWorkercs.Stopped;
            }
            catch (Exception ex)
            {
                _HashRunning = false;
                this.StatusWorker = EnStatusWorkercs.FatalError;

                if (this.onEvWorkerBaseError != null)
                {
                    EvWorkerBaseError ArgErrorW = new EvWorkerBaseError(this, ex.Message);
                    this.onEvWorkerBaseError.Invoke(this, ArgErrorW);
                }

                throw new ApplicationException(string.Format("{0}.Run   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
            }
        }


        /// <summary>
        /// Базовый класс для нашего пулакоторый представляет из себя список но управление списком будет осуществлять ребёнок этого класса
        /// </summary>
        public abstract class WorkerListBase : IEnumerable
        {
            /// <summary>
            /// Внутренний список 
            /// </summary>
            private List<WorkerBase> WrkL = new List<WorkerBase>();

            /// <summary>
            /// Количчество объектов в контейнере
            /// </summary>
            public int Count
            {
                get
                {
                    try
                    {
                        int rez;
                        lock (WrkL)
                        {
                            rez = WrkL.Count;
                        }
                        return rez;
                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException(string.Format("{0}.Count   Упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
                    }
                }
                private set { }
            }

            /// <summary>
            /// Добавление нового элемента
            /// </summary>
            /// <param name="newWrk">Элемент который нужно добавить в список</param>
            /// <param name="HashExeption">C отображением исключений</param>
            /// <returns>Результат операции (Успех или нет)</returns>
            protected bool Add(WorkerBase newWrk, bool HashExeption)
            {
                bool rez = false;

                try
                {
                    lock (this.WrkL)
                    {
                        newWrk.Index = WrkL.Count;
                        this.WrkL.Add(newWrk);
                        rez = true;
                    }
                }
                catch (Exception ex)
                {
                    if (HashExeption) throw new ApplicationException(string.Format("{0}.Add   Упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
                }
                return rez;
            }

            /// <summary>
            /// Удаление элемента
            /// </summary>
            /// <param name="delWrk">Элемент который нужно удалить из списка</param>
            /// <param name="HashExeption">C отображением исключений</param>
            /// <returns>Результат операции (Успех или нет)</returns>
            protected bool Remove(WorkerBase delWrk, bool HashExeption)
            {
                bool rez = false;
                try
                {
                    lock (this.WrkL)
                    {
                        int delIndex = delWrk.Index;
                        this.WrkL.RemoveAt(delIndex);

                        for (int i = delIndex; i < this.WrkL.Count; i++)
                        {
                            this.WrkL[i].Index = i;
                        }

                        rez = true;
                    }
                }
                catch (Exception ex)
                {
                    if (HashExeption) throw new ApplicationException(string.Format("Не удалось удалить элемент с мндексом {0} из списка. Произошла ошибка: {1}", delWrk.Index, ex.Message));
                }

                return rez;
            }


            /// <summary>
            /// Индексатор
            /// </summary>
            /// <param name="i">Индекс элемента в массиве</param>
            /// <returns>возвращаем объект</returns>
            public WorkerBase this[int i]
            {
                get
                {
                    try
                    {
                        WorkerBase rez = null;
                        lock (WrkL)
                        {
                            rez = this.WrkL[i];
                        }

                        if (rez == null) throw new ApplicationException(String.Format("Объект с индексом {0} не найден.", i));

                        return rez;
                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException(string.Format("{0}.getBookmarkComponent({1})   Упали с ошибкой: ({2})", this.GetType().Name, i, ex.Message));
                    }
                }
                private set { }
            }

            /// <summary>
            /// Для обращения по индексатору
            /// </summary>
            /// <returns>Возвращаем стандарнтый индексатор</returns>
            public IEnumerator GetEnumerator()
            {
                IEnumerator rez = null;
                lock (WrkL)
                {
                    rez = this.WrkL.GetEnumerator();
                }
                return rez;
            }
        }
    }
}
