using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDotx
{
    /// <summary>
    /// Класс представляет изсебя задание для сервера
    /// </summary>
    public class TaskExcel : Lib.TaskExcelBase
    {
        /// <summary>
        /// Если на основе задания был сделан результат то ссылка на этот результат присваивается в самом задании
        /// </summary>
        public new RezultTaskExcel RezTsk
        {
            get
            {
                return (RezultTaskExcel)base.RezTsk;
            }
        }

        /// <summary>
        /// Путь к файлу шаблона или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса
        /// </summary>
        public string Source { get; private set; }

        /// <summary>
        /// Путь к файлу отчёта который создать по окончании работы или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса
        /// </summary>
        public string Target { get; private set; }

        /// <summary>
        /// Замена в папке назначения файла если уже с таким именем файл существует
        /// </summary>
        public bool? ReplaseFileTarget { get; private set; }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="Source">Путь к файлу шаблона или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса</param>
        /// <param name="Target">Путь к файлу отчёта который создать по окончании работы или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса</param>
         /// <param name="TblL">Список таблиц который будем использовать</param>
        /// <param name="ReplaseFileTarget">Замена в папке назначения файла если уже ст таким именем файл существует</param>
        public TaskExcel(string Source, string Target, TableList TblL, bool? ReplaseFileTarget) : base(TblL)
        {
            try
            {
                this.Source = Source;
                this.Target = Target;
                this.ReplaseFileTarget = ReplaseFileTarget;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}   Упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
            }
        }
        //
        /// <summary>
        /// Процесс создания отчёта с подменой в шаблоне необходимых элементов на наши закладки и таблицы
        /// </summary>
        /// <param name="Source">Путь к файлу шаблона или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса</param>
        /// <param name="Target">Путь к файлу отчёта который создать по окончании работы или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса</param>
        /// <param name="Tbl">Таблица который будем использовать</param>
        /// <param name="ReplaseFileTarget">Замена в папке назначения файла если уже ст таким именем файл существует</param>
        public TaskExcel(string Source, string Target,  Table Tbl, bool ReplaseFileTarget) : this(Source, Target, new TableList(), ReplaseFileTarget)
        {
            try
            {
                TableList TblL = new TableList();
                TblL.Add(Tbl, true);
                base.setTableList(TblL);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.StartCreateReport   Упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
            }
        }
        //
        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="Source">Путь к файлу шаблона или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса</param>
        /// <param name="Target">Путь к файлу отчёта который создать по окончании работы или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса</param>
        /// <param name="TblL">Список таблиц который будем использовать</param>
        public TaskExcel(string Source, string Target, TableList TblL) : this(Source, Target, TblL, null)
        {
            try
            {

            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.StartCreateReport   Упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
            }
        }
        //
        /// <summary>
        /// Процесс создания отчёта с подменой в шаблоне необходимых элементов на наши закладки и таблицы
        /// </summary>
        /// <param name="Source">Путь к файлу шаблона или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса</param>
        /// <param name="Target">Путь к файлу отчёта который создать по окончании работы или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса</param>
        /// <param name="Tbl">Таблица который будем использовать</param>
        public TaskExcel(string Source, string Target, Table Tbl) : this(Source, Target,  new TableList(), null)
        {
            try
            {
                TableList TblL = new TableList();
                TblL.Add(Tbl, true);
                base.setTableList(TblL);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.StartCreateReport   Упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
            }
        }
    }
}
