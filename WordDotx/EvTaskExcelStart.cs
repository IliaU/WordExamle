using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDotx
{
    /// <summary>
    /// Аргументы события когда задание попало в работу
    /// </summary>
    public class EvTaskExcelStart : EventArgs
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
        /// Конструктор
        /// </summary>
        /// <param name="Tsk">Задание которое сейчас выполняется</param>
        /// <param name="ExlServ">Сервер который сейчас выполняет это задание</param>
        public EvTaskExcelStart(TaskExcel Tsk, ExcelServer ExlServ)
        {
            this.Tsk = Tsk;
            this.ExlServ = ExlServ;
        }
    }
}
