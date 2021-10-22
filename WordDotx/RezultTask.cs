using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDotx
{
    /// <summary>
    /// Класс который представляет результат работы нашего задания
    /// </summary>
    public class RezultTask : Lib.TaskWordBase.RezultTaskBase
    {
        /// <summary>
        /// Конструктор который позволяет при создании связать результат с заданием для того чтобы потом отслеживать его
        /// </summary>
        /// <param name="Tsk">Задание к которому нужно привязать создаваемый класс который будет иметь доступ к закрытым полям для прогресс бара</param>
        public RezultTask(TaskWord Tsk) : base(Tsk)
        {

        }
    }
}
