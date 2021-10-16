using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDotx
{
    /// <summary>
    /// Класс реализующий объект закладку для шаблона
    /// </summary>
    public class Bookmark: Lib.BookmarkBase
    {
        /// <summary>
        /// Значение закладки на которое нужно заменить это имя
        /// </summary>
        public string BookmarkValue { get; private set; }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="BookmarksName">Имя закладки которое нужно найти в файле шаблона</param>
        /// <param name="BookmarksValue">Значение закладки на которое нужно заменить это имя</param>
        public Bookmark(string BookmarksName, string BookmarkValue):base(BookmarksName)
        {
            try
            {
                this.BookmarkValue = BookmarkValue;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}   Упали с ошибкой в конструкторе: ({1})", this.GetType().Name, ex.Message));
            }
        }
    }
}
