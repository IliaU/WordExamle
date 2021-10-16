using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDotx
{
    /// <summary>
    /// Класс представляет из себя список наших объектов Bookmark
    /// </summary>
    public class BookmarkList:Lib.BookmarkBase.BookmarkListBase
    {
        /// <summary>
        /// Конструктор
        /// </summary>
        public BookmarkList()
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
