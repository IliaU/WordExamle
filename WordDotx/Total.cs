using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDotx
{
    /// <summary>
    /// Класс реализующий объект закладку для шаблона Total
    /// </summary>
    public class Total : Lib.TotalBase
    {
        /// <summary>
        /// Значение закладки на которое нужно заменить это имя
        /// </summary>
        public string TotalValue { get; private set; }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="TotalName">Имя закладки которое нужно найти в файле шаблона</param>
        /// <param name="TotalValue">Значение закладки на которое нужно заменить это имя</param>
        public Total(string TotalName, string TotalValue) :base(TotalName)
        {
            try
            {
                this.TotalValue = TotalValue;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}   Упали с ошибкой в конструкторе: ({1})", this.GetType().Name, ex.Message));
            }
        }
    }
}
