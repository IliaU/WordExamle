using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDotx
{
    /// <summary>
    /// Аргументы события исключения Работника
    /// </summary>
    public class EvWorkerExcelError : EventArgs
    {
        /// <summary>
        /// Работник который выполняет задание
        /// </summary>
        public WorkerExcel Exl { get; private set; }

        /// <summary>
        /// Сообщение об ошибке
        /// </summary>
        public string ErrorMessage { get; private set; }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="Exl">Работник который выполняет задание</param>
        /// <param name="ErrorMessage">Сообщение об ошибке</param>
        public EvWorkerExcelError(WorkerExcel Exl, string ErrorMessage)
        {
            this.Exl = Exl;
            this.ErrorMessage = ErrorMessage;
        }
    }
}
