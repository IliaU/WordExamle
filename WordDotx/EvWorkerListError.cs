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
    public class EvWorkerListError : EventArgs
    {
        /// <summary>
        /// Пул который выполняет задание
        /// </summary>
        public WorkerList WrkList { get; private set; }

        /// <summary>
        /// Сообщение об ошибке
        /// </summary>
        public string ErrorMessage { get; private set; }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="WrkList">Пул который выполняет задание</param>
        /// <param name="ErrorMessage">Сообщение об ошибке</param>
        public EvWorkerListError(WorkerList WrkList, string ErrorMessage)
        {
            this.WrkList = WrkList;
            this.ErrorMessage = ErrorMessage;
        }
    }
}
