using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDotx.Lib
{
    /// <summary>
    /// Базовый класс для заданий которые могут выполнять задания из очереди асинхронно
    /// </summary>
    public abstract class TaskWordBase
    {
        /// <summary>
        /// Идентификатор задания
        /// </summary>
        public Guid Sid { get; private set; }

        /// <summary>
        /// Если на основе задания был сделан результат то ссылка на этот результат присваивается в самом задании
        /// </summary>
        protected RezultTaskBase RezTsk;

        /// <summary>
        /// Статус задания
        /// </summary>
        public EnStatusTask StatusTask { get; private set; }

        /// <summary>
        /// Список сообщений из стека для того чтобы пользователь мог посмотреть ошибку если задание упало
        /// </summary>
        public List<string> StatusMessage { get; private set; }

        /// <summary>
        /// Список таблиц который будем использовать
        /// </summary>
        public TableList TblL { get; private set; }
        
        /// <summary>
        /// Реальное создание задания
        /// </summary>
        public DateTime CraeteDt { get; private set; }

        /// <summary>
        /// Реальное установка в очередь нашего задания
        /// </summary>
        public DateTime? PendingProcessing { get; private set; }

        /// <summary>
        /// Реальное начало формированиея отчёта
        /// </summary>
        public DateTime? StartProcessing { get; private set; }

        /// <summary>
        /// Реальное окончание формированиея отчёта
        /// </summary>
        public DateTime? EndProcessing { get; private set; }

        /// <summary>
        /// Конструктор
        /// </summary>
        public TaskWordBase(TableList TblL)
        {
            try
            {
                this.Sid = Guid.NewGuid();
                this.CraeteDt = DateTime.Now;
                this.StatusTask = EnStatusTask.None;
                this.StatusMessage = new List<string>();
                this.TblL = TblL;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}  При создании объекта в конструкторе упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Установка листа с таблицами
        /// </summary>
        /// <param name="TblL">Новый лист с таблицами</param>
        protected void setTableList(TableList TblL)
        {
            try
            {
                this.TblL = TblL;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}  При установке листа с таблицами упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Базовый класс который представляет из себя результат
        /// </summary>
        public abstract class RezultTaskBase
        {
            /// <summary>
            /// Задание которое выполняется и за результатом которого мы следим
            /// </summary>
            public TaskWordBase Tsk { get; private set; }

            /// <summary>
            /// Список по каждой таблице внутри шаблона (их может быть больше чем в источнике и количество строк которое уже в каждой из них залито)
            /// </summary>
            public List<RezultTaskAffectetdRow> TableInWordAffectedRowList { get; private set; }

            /// <summary>
            /// Конструктор который позволяет при создании связать результат с заданием для того чтобы потом отслеживать его
            /// </summary>
            /// <param name="Tsk">Задание к которому нужно привязать создаваемый класс который будет иметь доступ к закрытым полям для прогресс бара</param>
            public RezultTaskBase(TaskWordBase Tsk)
            {
                try
                {
                    // Присвоить ссылки на друг друга заданию и резултату чтобы они могли работать друг с другом
                    this.Tsk = Tsk;
                    Tsk.RezTsk = this;

                    // Создаём пустой список с таблицами по которым будем выводить статистику из ворда
                    this.TableInWordAffectedRowList = new List<RezultTaskAffectetdRow>();

                    // Если мы хотим получить класс который наблюдает за результатом значит мы будем выполнять асинхронно, тогда пока задания нет в очереди мы просто поменяем статус
                    this.Tsk.StatusTask = EnStatusTask.Empty;
                }
                catch (Exception ex)
                {
                    throw new ApplicationException(string.Format("{0}  При создании объекта в конструкторе упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
                }
            }

            /// <summary>
            /// Базовый класс для сервера чтобы он мог влиять на поля задания и прогресс бар задания
            /// </summary>
            public abstract class WordDotxServerBase
            {
                /// <summary>
                /// Конструктор
                /// </summary>
                public WordDotxServerBase()
                {
                    try
                    {

                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException(string.Format("{0}  При создании объекта в конструкторе упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
                    }
                }

                /// <summary>
                /// Устанавливает статус когда задача поподает в очередь
                /// </summary>
                /// <param name="Tsk">Задание которое попадает в очередь</param>
                /// <param name="Stat">Статус который надо выставить</param>
                protected void SetStatusTaskWord(TaskWordBase Tsk, EnStatusTask Stat)
                {
                    try
                    {
                        switch (Stat)
                        {
                            case EnStatusTask.Running:
                                // Устанавливаем реальный старт процесса формирования отчёта
                                Tsk.StartProcessing = DateTime.Now;
                                break;
                            case EnStatusTask.Success:
                            case EnStatusTask.ERROR:
                                // Устанавливаем реальное окончание процесса формирования отчёта
                                Tsk.EndProcessing = DateTime.Now;
                                break;
                            default:
                                break;
                        }

                        Tsk.StatusTask = Stat;
                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException(string.Format("{0}.SetStatusTaskWord  При установки нового статуса упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
                    }
                }

                /// <summary>
                /// Запись в результат системного сообщения чтобы перезать клиенту если он получает результат асинхронно
                /// </summary>
                /// <param name="Tsk">Задание которое обрабатываем</param>
                /// <param name="mes">Сообщение которое нужно добавить</param>
                protected void SetStatusMessage(TaskWordBase Tsk, string mes)
                {
                    try
                    {
                        Tsk.StatusMessage.Add(mes);
                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException(string.Format("{0}.SetStatusMessage  При установки нового статуса упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
                    }
                }

                /// <summary>
                /// Добавление в статистику имнформации по первоначальной обработке нашей таблицы из ворда
                /// </summary>
                /// <param name="Tsk">Задание в рамках которого происходит добавление задания в лист</param>
                /// <param name="StatTblRow">Создаём объект который потом будем просматривать</param>
                protected void SetInitTableInWordAffected(TaskWordBase Tsk, RezultTaskAffectetdRow StatTblRow)
                {
                    try
                    {
                        if (StatTblRow != null)
                        {
                            if (Tsk != null && Tsk.RezTsk != null && Tsk.RezTsk.TableInWordAffectedRowList != null)
                            {
                                lock (Tsk.RezTsk.TableInWordAffectedRowList)
                                {
                                    Tsk.RezTsk.TableInWordAffectedRowList.Add(StatTblRow);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException(string.Format("{0}.SetStatusMessage  При установки нового статуса упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
                    }
                }

                /// <summary>
                /// Добавление в статистику имнформации по первоначальной обработке нашей таблицы из ворда
                /// </summary>
                /// <param name="Tsk">Задание в рамках которого происходит добавление задания в лист</param>
                /// <param name="AffectedRow">Устанавливаем текущее кол-во строк которое уже обработали</param>
                protected void SetTableInWordAffected(TaskWordBase Tsk, int AffectedRow)
                {
                    try
                    {
                        if (Tsk != null && Tsk.RezTsk != null && Tsk.RezTsk.TableInWordAffectedRowList != null && Tsk.RezTsk.TableInWordAffectedRowList.Count > 0)
                        {
                            lock (Tsk.RezTsk.TableInWordAffectedRowList)
                            {
                                int FlagMaxIndex = 0;
                                for (int i = 0; i < Tsk.RezTsk.TableInWordAffectedRowList.Count; i++)
                                {
                                    if (!Tsk.RezTsk.TableInWordAffectedRowList[i].HashEnd) FlagMaxIndex = i;
                                }

                                Tsk.RezTsk.TableInWordAffectedRowList[FlagMaxIndex].AffectedRow = AffectedRow;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException(string.Format("{0}.SetStatusMessage  При установки нового статуса упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
                    }
                }


                /// <summary>
                /// Добавление в статистику имнформации по первоначальной обработке нашей таблицы из ворда
                /// </summary>
                /// <param name="Tsk">Задание в рамках которого происходит добавление задания в лист</param>
                /// <param name="AffectedRow">Устанавливаем текущее кол-во строк которое уже обработали</param>
                protected void SetEndTableInWordAffected(TaskWordBase Tsk, int AffectedRow)
                {
                    try
                    {
                        if (Tsk != null && Tsk.RezTsk != null && Tsk.RezTsk.TableInWordAffectedRowList != null && Tsk.RezTsk.TableInWordAffectedRowList.Count > 0)
                        {
                            lock (Tsk.RezTsk.TableInWordAffectedRowList)
                            {
                                int FlagMaxIndex = 0;
                                for (int i = 0; i < Tsk.RezTsk.TableInWordAffectedRowList.Count; i++)
                                {
                                    if (!Tsk.RezTsk.TableInWordAffectedRowList[i].HashEnd) FlagMaxIndex = i;
                                }

                                Tsk.RezTsk.TableInWordAffectedRowList[FlagMaxIndex].AffectedRow = AffectedRow;
                                Tsk.RezTsk.TableInWordAffectedRowList[FlagMaxIndex].HashEnd = true;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException(string.Format("{0}.SetStatusMessage  При установки нового статуса упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
                    }
                }
            }
        }

        /// <summary>
        /// Базовый класс для фермы чтобы мог менять статусу у нашего задания он же создаёт структуру очереди для потока который внутри
        /// </summary>
        public abstract class FarmWordDotxBase
        {

            /// <summary>
            /// Очередь для наших документов которые будут обрабатываться нашей фермой
            /// </summary>
            private static Queue<TaskWord> _QueTaskWord = new Queue<TaskWord>();

            /// <summary>
            /// Кол-во объектов в очереди 
            /// </summary>
            public static int QueTaskWordCount
            {
                get
                {
                    int rez = 0;
                    lock (_QueTaskWord)
                    {
                        rez = _QueTaskWord.Count;
                    }
                    return rez;
                }
                private set { }
            }

            /// <summary>
            /// Тайм аут между циклами воркера если заданий болльше нет;
            /// </summary>
            public static int TimeoutForWorkerSec = 3;

            /// <summary>
            /// Добвление в очередь элемена который потом нужно выполнить нашему пулу
            /// </summary>
            /// <param name="Tsk">Задание которое нужно выполнить нашему серверу</param>
            /// <returns>едоставляет класс через который пользрватель сможет наблюдать за состоянием нашего задания</returns>
            public static RezultTask QueTaskWordAdd(TaskWord Tsk)
            {
                try
                {
                    // Создаём класс через который пользователь будет наблюдать за выполнением задания и на который он сможет если надо подписаться
                    RezultTask rez = new RezultTask(Tsk);

                    // Устанавливаем время в которое поместили в очередь наше задание
                    Tsk.PendingProcessing = DateTime.Now;

                    lock (_QueTaskWord)
                    {
                        _QueTaskWord.Enqueue(Tsk);

                        // Выставляем статус для того чтобы пользователь мог увидеть что задание уже в очереди
                        Tsk.StatusTask = EnStatusTask.Pending;
                    }

                    // Возвращаемобъект пользователю
                    return rez;
                }
                catch (Exception ex)
                {
                    throw new ApplicationException(string.Format("{0}.QueTaskWordAdd   Упали с ошибкой при добавлении в очередь: ({1})", "FarmWordDotx", ex.Message));
                }
            }

            /// <summary>
            /// Возвращет объект из очереди но не всем а только объетум Worker
            /// </summary>
            /// <returns>Задание которое стоит в очереди</returns>
            private static TaskWord QueTaskWordGet()
            {
                try
                {
                    if (_QueTaskWord == null) throw new ApplicationException("Не инициирован FARM по этому обраблотка асинхронная не возможна");

                    TaskWord rez = null;

                    lock (_QueTaskWord)
                    {
                        if (_QueTaskWord.Count > 0) rez = _QueTaskWord.Dequeue();
                    }

                    // Возвращаемобъект пользователю
                    return rez;
                }
                catch (Exception ex)
                {
                    throw new ApplicationException(string.Format("{0}.QueTaskWordGet   Упали с ошибкой при извлечении объекта из очереди в очередь: ({1})", "FarmWordDotx", ex.Message));
                }
            }

            /// <summary>
            /// Класс через который воркер сможет менять статус и добавлять описание ошибки в стек и будет иметь доступ к очереди в потоке
            /// </summary>
            public abstract class WorkerBaseInclude
            {
                /// <summary>
                /// Получить задание из очереди
                /// </summary>
                /// <returns></returns>
                protected TaskWord QueTaskWordGet()
                {
                    try
                    {
                        // Возвращаемобъект пользователю
                        return FarmWordDotx.QueTaskWordGet();
                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException(string.Format("{0}.QueTaskWordGet   Упали с ошибкой при извлечении объекта из очереди в очередь: ({1})", "FarmWordDotx", ex.Message));
                    }
                }

                /// <summary>
                /// Получить ссылку на объект с результаттом данного задания чтобы можно было его править
                /// </summary>
                /// <param name="Tsk"></param>
                /// <returns></returns>
                protected RezultTaskBase GetRezult(TaskWord Tsk)
                {
                    try
                    {
                        // Возвращаемобъект пользователю
                        if (Tsk != null) return Tsk.RezTsk;
                        return null;

                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException(string.Format("{0}.GetRezult   Упали с ошибкой при извлечении объекта из очереди в очередь: ({1})", "FarmWordDotx", ex.Message));
                    }
                }

                /// <summary>
                /// Устанавливает статус когда задача поподает в очередь
                /// </summary>
                /// <param name="Tsk">Задание которое попадает в очередь</param>
                /// <param name="Stat">Статус который надо выставить</param>
                protected void SetStatusTaskWord(TaskWordBase Tsk, EnStatusTask Stat)
                {
                    try
                    {
                        if (Tsk != null) Tsk.StatusTask = Stat;
                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException(string.Format("{0}.SetStatusTaskWord  При установки нового статуса упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
                    }
                }

                /// <summary>
                /// Запись в результат системного сообщения чтобы перезать клиенту если он получает результат асинхронно
                /// </summary>
                /// <param name="Tsk">Задание которое обрабатываем</param>
                /// <param name="mes">Сообщение которое нужно добавить</param>
                protected void SetStatusMessage(TaskWordBase Tsk, string mes)
                {
                    try
                    {
                        Tsk.StatusMessage.Add(mes);
                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException(string.Format("{0}.SetStatusMessage  При установки нового статуса упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
                    }
                }
            }
        }
    }
}
