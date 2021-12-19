using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDotx
{
    /// <summary>
    /// Статус в которых может находиться задание
    /// </summary>
    public enum EnStatusTask
    {
        /// <summary>
        /// Задание не в очереди и не выполняется асинхронно
        /// </summary>
        None,
        /// <summary>
        /// Задание ещё не поставлено в очередь но уже имеет результат значит скорее всего пользователь потом запустит асинхронно
        /// </summary>
        Empty,
        /// <summary>
        /// Ожидание в очереди
        /// </summary>
        Pending,
        /// <summary>
        /// Worker взял в работу это задание
        /// </summary>
        Start,
        /// <summary>
        /// Задание начало считаться
        /// </summary>
        Running,
        /// <summary>
        /// Вызвано обновление дангных в документе
        /// </summary>
        Refresh,
        /// <summary>
        /// Задание выполнено и в процессе сохранения на диск
        /// </summary>
        Save,
        /// <summary>
        /// Задание выполнено
        /// </summary>
        Success,
        /// <summary>
        /// Выполнено частично
        /// </summary>
        WARNING,
        /// <summary>
        /// Ошибка при выполнении
        /// </summary>
        ERROR
    }
}
