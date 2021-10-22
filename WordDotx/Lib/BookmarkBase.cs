using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Collections;

namespace WordDotx.Lib
{
    /// <summary>
    /// Базовый класс для закладок
    /// </summary>
    public abstract class BookmarkBase
    {
        /// <summary>
        /// Индекс элемента в списке
        /// </summary>
        public int Index { get; private set; }

        /// <summary>
        /// Имя закладки которое нужно найти в файле шаблона
        /// </summary>
        public string BookmarkName { get; private set; }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="BookmarkName">Имя закладки</param>
        public BookmarkBase(string BookmarkName)
        {
            try
            {
                this.Index = -1;
                this.BookmarkName = BookmarkName;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}   Упали с ошибкой в конструкторе: ({1})", this.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Базовый класс для компонента списка эелементов Bookmark
        /// </summary>
        public abstract class BookmarkListBase : IEnumerable
        {
            /// <summary>
            /// Внутренний список 
            /// </summary>
            private List<BookmarkBase> BkmL = new List<BookmarkBase>();

            /// <summary>
            /// Количчество объектов в контейнере
            /// </summary>
            public int Count
            {
                get
                {
                    try
                    {
                        int rez;
                        lock (BkmL)
                        {
                            rez = BkmL.Count;
                        }
                        return rez;
                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException(string.Format("{0}.Count   Упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
                    }
                }
                private set { }
            }

            /// <summary>
            /// Добавление нового элемента
            /// </summary>
            /// <param name="newBkm">Элемент который нужно добавить в список</param>
            /// <param name="HashExeption">C отображением исключений</param>
            /// <returns>Результат операции (Успех или нет)</returns>
            public bool Add(BookmarkBase newBkm, bool HashExeption)
            {
                bool rez = false;

                try
                {
                    lock (this.BkmL)
                    {
                        // Проверка на наличие этого элемента в списке
                        foreach (BookmarkBase item in this.BkmL)
                        {
                            if (item.BookmarkName == newBkm.BookmarkName)
                            {
                                throw new ApplicationException(string.Format("Элемент с таким именем: {0} уже существует в списке.", newBkm.BookmarkName));
                            }
                        }

                        newBkm.Index = BkmL.Count;
                        this.BkmL.Add(newBkm);
                        rez = true;
                    }
                }
                catch (Exception ex)
                {
                    if(HashExeption) throw new ApplicationException(string.Format("{0}.Add   Упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
                }
                return rez;
            }

            /// <summary>
            /// Удаление элемента
            /// </summary>
            /// <param name="delBkm">Элемент который нужно удалить из списка</param>
            /// <param name="HashExeption">C отображением исключений</param>
            /// <returns>Результат операции (Успех или нет)</returns>
            public bool Remove(BookmarkBase delBkm, bool HashExeption)
            {
                bool rez = false;
                try
                {
                    lock (this.BkmL)
                    {
                        int delIndex = delBkm.Index;
                        this.BkmL.RemoveAt(delIndex);

                        for (int i = delIndex; i < this.BkmL.Count; i++)
                        {
                            this.BkmL[i].Index = i;
                        }

                        rez = true;
                    }
                }
                catch (Exception ex)
                {
                    if (HashExeption) throw new ApplicationException(string.Format("Не удалось удалить элемент с именем {0} из списка. Произошла ошибка: {1}", delBkm.BookmarkName, ex.Message));
                }

                return rez;
            }

            /// <summary>
            /// Обновление данных элемента конфигурации.
            /// </summary>
            /// <param name="IndexId">Индекс элемента который нужно обновить</param>
            /// <param name="updBkm">Пользователь у которого нужно изменить данные</param>
            /// <param name="HashExeption">C отображением исключений</param>
            /// <returns>Результат операции (Успех или нет)</returns>
            public bool Update(int IndexId, BookmarkBase updBkm, bool HashExeption)
            {
                bool rez = false;
                try
                {
                    lock (this.BkmL)
                    {

                        if (IndexId >= this.BkmL.Count)
                        {
                            if (HashExeption) throw new ApplicationException(string.Format("Не удалось обновить данные элемента в списке {0}. Элемента с таким индексом {1} не существует.", updBkm.BookmarkName, updBkm.ToString()));
                        }
                        else
                        {
                            updBkm.Index = IndexId;
                            this.BkmL[IndexId] = updBkm;

                            rez = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (HashExeption) throw new ApplicationException(string.Format("Не удалось обновить данные элемента в списке {0}. Произошла ошибка: {1}", updBkm.BookmarkName, ex.Message));
                }

                return rez;
            }


            /// <summary>
            /// Индексатор
            /// </summary>
            /// <param name="i">Индекс элемента в массиве</param>
            /// <returns>возвращаем объект</returns>
            public BookmarkBase this[int i]
            {
                get
                {
                    try
                    {
                        BookmarkBase rez = null;
                        lock (BkmL)
                        {
                            rez = this.BkmL[i];
                        }

                        if (rez == null) throw new ApplicationException(String.Format("Объект с индексом {0} не найден.", i));

                        return rez;
                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException(string.Format("{0}.getBookmarkComponent({1})   Упали с ошибкой: ({2})", this.GetType().Name, i, ex.Message));
                    }
                }
                private set { }
            }


            /// <summary>
            /// Индексатор
            /// </summary>
            /// <param name="s">Введите имя закладки</param>
            /// <returns>возвращаем объект</returns>
            public BookmarkBase this[string s]
            {
                get
                {
                    try
                    {
                        BookmarkBase rez = null;
                        lock (BkmL)
                        {
                            foreach (BookmarkBase item in this.BkmL)
                            {
                                if (item.BookmarkName == s)
                                {
                                    rez = item;
                                    break;
                                }
                            }
                        }

                        if (rez == null) throw new ApplicationException(String.Format("Объект с именем {0} не найден.", s));

                        return rez;
                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException(string.Format("{0}.getBookmarkComponent({1})   Упали с ошибкой: ({2})", this.GetType().Name, s, ex.Message));
                    }
                }
                private set { }
            }

            /// <summary>
            /// Для обращения по индексатору
            /// </summary>
            /// <returns>Возвращаем стандарнтый индексатор</returns>
            public IEnumerator GetEnumerator()
            {
                IEnumerator rez = null;
                lock (BkmL)
                {
                    rez = this.BkmL.GetEnumerator();
                }
                return rez;
            }
        }
    }
}
