using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;

namespace WordDotx
{
    /// <summary>
    /// Реализация таблицы с которой потом работаем в шаблоне ворда
    /// </summary>
    public class Table : Lib.TableBase
    {
        /// <summary>
        /// Сама таблица которую будем использовать при подстановке например когд у нас указано в таблице значение @D<Индекс> или @D<TableName>
        /// </summary>
        public DataTable TableValue;

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="TableName">Имя таблицы с которой работаем</param>
        /// <param name="TableValue">Сама таблица которую будем использовать при подстановке например когд у нас указано в таблице значение @D<Индекс> или @D<TableName></param>
        public Table(string TableName, DataTable TableValue) :base(TableName)
        {
            try
            {
                this.TableValue = TableValue;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}   Упали с ошибкой в конструкторе: ({1})", this.GetType().Name, ex.Message));
            }
        }
    }
}
