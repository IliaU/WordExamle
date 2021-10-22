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
    public class EvTaskWordError : EventArgs
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
        /// Сообщение об ошибке
        /// </summary>
        public string ErrorMessage { get; private set; }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="Tsk">Задание которое сейчас выполняется</param>
        /// <param name="WordServ">Сервер который сейчас выполняет это задание</param>
        /// <param name="ErrorMessage">Сообщение об ошибке</param>
        public EvTaskWordError(TaskWord Tsk, WordDotxServer WordServ, string ErrorMessage)
        {
            this.Tsk = Tsk;
            this.WordServ = WordServ;
            this.ErrorMessage = ErrorMessage;
        }
    }
}
