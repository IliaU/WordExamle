﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDotx
{
    /// <summary>
    /// Аргументы события исключения Работника
    /// </summary>
    public class EvWorkerWordError : EventArgs
    {
        /// <summary>
        /// Работник который выполняет задание
        /// </summary>
        public WorkerWord Wrk { get; private set; }

        /// <summary>
        /// Сообщение об ошибке
        /// </summary>
        public string ErrorMessage { get; private set; }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="Wrk">Работник который выполняет задание</param>
        /// <param name="ErrorMessage">Сообщение об ошибке</param>
        public EvWorkerWordError(WorkerWord Wrk, string ErrorMessage)
        {
            this.Wrk = Wrk;
            this.ErrorMessage = ErrorMessage;
        }
    }
}
