using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDotx
{
    /// <summary>
    /// Аргументы события исключения пула
    /// </summary>
    public class EvWorkerExcelListError : EventArgs
    {
        /// <summary>
        /// Пул который выполняет задание
        /// </summary>
        public WorkerExcelList ExlList { get; private set; }

        /// <summary>
        /// Сообщение об ошибке
        /// </summary>
        public string ErrorMessage { get; private set; }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="ExlList">Пул который выполняет задание</param>
        /// <param name="ErrorMessage">Сообщение об ошибке</param>
        public EvWorkerExcelListError(WorkerExcelList ExlList, string ErrorMessage)
        {
            this.ExlList = ExlList;
            this.ErrorMessage = ErrorMessage;
        }
    }
}
