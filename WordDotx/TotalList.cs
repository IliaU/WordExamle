using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDotx
{
    public class TotalList : Lib.TotalBase.TotalListBase
    {
        /// <summary>
        /// Конструктор
        /// </summary>
        public TotalList()
        {
            try
            {

            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}   Упали с ошибкой в конструкторе: ({1})", this.GetType().Name, ex.Message));
            }
        }
    }
}
