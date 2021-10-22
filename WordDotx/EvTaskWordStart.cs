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
    public class EvTaskWordStart : EventArgs
    {
        /// <summary>
        /// Задание которое сейчас выполняется
        /// </summary>
        public TaskWord Tsk { get; private set; }

        /// <summary>
        /// Сервер который сейчас выполняет это задание
        /// </summary>
        public WordDotxServer WordServ { get; private set; }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="Tsk">Задание которое сейчас выполняется</param>
        /// <param name="WordServ">Сервер который сейчас выполняет это задание</param>
        public EvTaskWordStart(TaskWord Tsk, WordDotxServer WordServ)
        {
            this.Tsk = Tsk;
            this.WordServ = WordServ;
        }
    }
}
