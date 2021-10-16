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
    public abstract class TotalBase
    {
        /// <summary>
        /// Индекс элемента в списке
        /// </summary>
        public int Index { get; private set; }

        /// <summary>
        /// Имя тотала которое нужно найти в файле шаблона
        /// </summary>
        public string TotalName { get; private set; }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="TotalName">Имя тотала</param>
        public TotalBase(string TotalName)
        {
            try
            {
                this.TotalName = TotalName;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}   Упали с ошибкой в конструкторе: ({1})", this.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Базовый класс для компонента списка эелементов Total
        /// </summary>
        public abstract class TotalListBase : IEnumerable
        {
            /// <summary>
            /// Внутренний список 
            /// </summary>
            private List<TotalBase> TtlL = new List<TotalBase>();

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
                        lock (TtlL)
                        {
                            rez = TtlL.Count;
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
            /// <param name="newTtl">Элемент который нужно добавить в список</param>
            /// <param name="HashExeption">C отображением исключений</param>
            /// <returns>Результат операции (Успех или нет)</returns>
            public bool Add(TotalBase newTtl, bool HashExeption)
            {
                bool rez = false;

                try
                {
                    lock (this.TtlL)
                    {
                        // Проверка на наличие этого элемента в списке
                        foreach (TotalBase item in this.TtlL)
                        {
                            if (item.TotalName == newTtl.TotalName)
                            {
                                throw new ApplicationException(string.Format("Элемент с таким именем: {0} уже существует в списке.", newTtl.TotalName));
                            }
                        }

                        newTtl.Index = TtlL.Count;
                        this.TtlL.Add(newTtl);
                        rez = true;
                    }
                }
                catch (Exception ex)
                {
                    if (HashExeption) throw new ApplicationException(string.Format("{0}.Add   Упали с ошибкой: ({1})", this.GetType().Name, ex.Message));
                }
                return rez;
            }

            /// <summary>
            /// Удаление элемента
            /// </summary>
            /// <param name="delTtl">Элемент который нужно удалить из списка</param>
            /// <param name="HashExeption">C отображением исключений</param>
            /// <returns>Результат операции (Успех или нет)</returns>
            public bool Remove(TotalBase delTtl, bool HashExeption)
            {
                bool rez = false;
                try
                {
                    lock (this.TtlL)
                    {
                        int delIndex = delTtl.Index;
                        this.TtlL.RemoveAt(delIndex);

                        for (int i = delIndex; i < this.TtlL.Count; i++)
                        {
                            this.TtlL[i].Index = i;
                        }

                        rez = true;
                    }
                }
                catch (Exception ex)
                {
                    if (HashExeption) throw new ApplicationException(string.Format("Не удалось удалить элемент с именем {0} из списка. Произошла ошибка: {1}", delTtl.TotalName, ex.Message));
                }

                return rez;
            }

            /// <summary>
            /// Обновление данных элемента конфигурации.
            /// </summary>
            /// <param name="IndexId">Индекс элемента который нужно обновить</param>
            /// <param name="updTtl">Пользователь у которого нужно изменить данные</param>
            /// <param name="HashExeption">C отображением исключений</param>
            /// <returns>Результат операции (Успех или нет)</returns>
            public bool Update(int IndexId, TotalBase updTtl, bool HashExeption)
            {
                bool rez = false;
                try
                {
                    lock (this.TtlL)
                    {

                        if (IndexId >= this.TtlL.Count)
                        {
                            if (HashExeption) throw new ApplicationException(string.Format("Не удалось обновить данные элемента в списке {0}. Элемента с таким индексом {1} не существует.", updTtl.TotalName, updTtl.ToString()));
                        }
                        else
                        {
                            updTtl.Index = IndexId;
                            this.TtlL[IndexId] = updTtl;

                            rez = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (HashExeption) throw new ApplicationException(string.Format("Не удалось обновить данные элемента в списке {0}. Произошла ошибка: {1}", updTtl.TotalName, ex.Message));
                }

                return rez;
            }

            /// <summary>
            /// Получение компонента по его ID
            /// </summary>
            /// <param name="i">Введите идентификатор</param>
            /// <returns></returns>
            public TotalBase getTotalComponent(int i)
            {
                try
                {
                    TotalBase rez = null;
                    lock (TtlL)
                    {
                        rez = this.TtlL[i];
                    }

                    if (rez == null) throw new ApplicationException(String.Format("Объект с индексом {0} не найден.", i));

                    return rez;
                }
                catch (Exception ex)
                {
                    throw new ApplicationException(string.Format("{0}.getBookmarkComponent({1})   Упали с ошибкой: ({2})", this.GetType().Name, i, ex.Message));
                }
            }

            /// <summary>
            /// Получение компонента по его имени
            /// </summary>
            /// <param name="s">Введите имя закладки</param>
            /// <returns></returns>
            public TotalBase getTotalComponent(string s)
            {
                try
                {
                    TotalBase rez = null;
                    lock (TtlL)
                    {
                        foreach (TotalBase item in this.TtlL)
                        {
                            if (item.TotalName == s)
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

            /// <summary>
            /// Для обращения по индексатору
            /// </summary>
            /// <returns>Возвращаем стандарнтый индексатор</returns>
            public IEnumerator GetEnumerator()
            {
                IEnumerator rez = null;
                lock (TtlL)
                {
                    rez = this.TtlL.GetEnumerator();
                }
                return rez;
            }
        }
    }
}
