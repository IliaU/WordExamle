using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDotx
{
    /// <summary>
    /// Статусы в которых может находиться работник
    /// </summary>
    public enum EnStatusWorkercs
    {
        /// <summary>
        /// Создан но не запущен
        /// </summary>
        Created,
        /// <summary>
        /// Выполняет какое то задание
        /// </summary>
        Running,
        /// <summary>
        /// Ожидает задание для выполнения
        /// </summary>
        Waiting,
        /// <summary>
        /// В проуессе остановки
        /// </summary>
        Stopping,
        /// <summary>
        /// Остановлено
        /// </summary>
        Stopped,
        /// <summary>
        /// Фатальная ошибка
        /// </summary>
        FatalError
    }
}
