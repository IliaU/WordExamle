using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDotx
{
    /// <summary>
    /// Аргументы события когда задание упало с ошибкой
    /// </summary>
    public class EvTaskExcelError : EventArgs
    {
        /// <summary>
        /// Задание которое сейчас выполняется
        /// </summary>
        public TaskExcel Tsk { get; private set; }

        /// <summary>
        /// Сервер который сейчас выполняет это задание
        /// </summary>
        public ExcelServer ExlServ { get; private set; }

        /// <summary>
        /// Сообщение об ошибке
        /// </summary>
        public string ErrorMessage { get; private set; }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="Tsk">Задание которое сейчас выполняется</param>
        /// <param name="WordServ">Сервер который сейчас выполняет это задание</param>
        /// <param name="ErrorMessage">Сообщение об ошибке</param>
        public EvTaskExcelError(TaskExcel Tsk, ExcelServer ExlServ, string ErrorMessage)
        {
            this.Tsk = Tsk;
            this.ExlServ = ExlServ;
            this.ErrorMessage = ErrorMessage;
        }
    }
}
