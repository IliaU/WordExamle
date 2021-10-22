using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDotx
{
    /// <summary>
    /// класс который сопоставляет конкретную таблицу на основе которой строится статистика и то кол-во строк которое обработано в потоке по этой таблице
    /// </summary>
    public class RezultTaskAffectetdRow
    {
        /// <summary>
        /// Таблица на основе которой строится статистика из базы данных
        /// </summary>
        public Table Tbl;

        /// <summary>
        /// Сколько строк уже обработано в этой таблице
        /// </summary>
        public int AffectedRow;

        /// <summary>
        /// Достигли конца по обработке этой таблицы
        /// </summary>
        public bool HashEnd;

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="Tbl">Таблица на основе которой идёт построение отчёта</param>
        public RezultTaskAffectetdRow(Table Tbl)
        {
            try
            {
                this.Tbl = Tbl;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}  При создании объекта в конструкторе упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
            }
        }
    }
}
