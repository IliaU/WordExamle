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
    public class EvWorkerWordListError : EventArgs
    {
        /// <summary>
        /// Пул который выполняет задание
        /// </summary>
        public WorkerWordList WrkList { get; private set; }

        /// <summary>
        /// Сообщение об ошибке
        /// </summary>
        public string ErrorMessage { get; private set; }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="WrkList">Пул который выполняет задание</param>
        /// <param name="ErrorMessage">Сообщение об ошибке</param>
        public EvWorkerWordListError(WorkerWordList WrkList, string ErrorMessage)
        {
            this.WrkList = WrkList;
            this.ErrorMessage = ErrorMessage;
        }
    }
}
