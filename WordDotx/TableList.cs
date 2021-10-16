using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDotx
{
    /// <summary>
    /// Класс представляет из себя список наших таблиц
    /// </summary>
    public class TableList : Lib.TableBase.TableListBase
    {
        /// <summary>
         /// Конструктор
         /// </summary>
        public TableList()
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
