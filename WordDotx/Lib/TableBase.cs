using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Collections;
using System.Data;

namespace WordDotx.Lib
{
    /// <summary>
    /// Базовый класс для таблиц с которыми работаем
    /// </summary>
    public abstract class TableBase
    {
        /// <summary>
        /// Индекс элемента в списке
        /// </summary>
        public int Index { get; private set; }

        /// <summary>
        /// Имя закладки которое нужно найти в файле шаблона
        /// </summary>
        public string TableName { get; private set; }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="TableName">Имя таблицы</param>
        public TableBase(string TableName)
        {
            try
            {
                this.TableName = TableName;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}   Упали с ошибкой в конструкторе: ({1})", this.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Базовый класс для компонента списка эелементов конфигурации
        /// </summary>
        public abstract class TableListBase : IEnumerable
        {
            /// <summary>
            /// Внутренний список 
            /// </summary>
            private List<TableBase> TblL = new List<TableBase>();

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
                        lock (TblL)
                        {
                            rez = TblL.Count;
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
            /// <param name="newTbl">Элемент который нужно добавить в список</param>
            /// <param name="HashExeption">C отображением исключений</param>
            /// <returns>Результат операции (Успех или нет)</returns>
            public bool Add(TableBase newTbl, bool HashExeption)
            {
                bool rez = false;

                try
                {
                    lock (this.TblL)
                    {
                        // Проверка на наличие этого элемента в списке
                        foreach (TableBase item in this.TblL)
                        {
                            if (item.TableName == newTbl.TableName)
                            {
                                throw new ApplicationException(string.Format("Элемент с таким именем: {0} уже существует в списке.", newTbl.TableName));
                            }
                        }

                        newTbl.Index = TblL.Count;
                        this.TblL.Add(newTbl);
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
            /// <param name="delTbl">Элемент который нужно удалить из списка</param>
            /// <param name="HashExeption">C отображением исключений</param>
            /// <returns>Результат операции (Успех или нет)</returns>
            public bool Remove(TableBase delTbl, bool HashExeption)
            {
                bool rez = false;
                try
                {
                    lock (this.TblL)
                    {
                        int delIndex = delTbl.Index;
                        this.TblL.RemoveAt(delIndex);

                        for (int i = delIndex; i < this.TblL.Count; i++)
                        {
                            this.TblL[i].Index = i;
                        }

                        rez = true;
                    }
                }
                catch (Exception ex)
                {
                    if (HashExeption) throw new ApplicationException(string.Format("Не удалось удалить элемент с именем {0} из списка. Произошла ошибка: {1}", delTbl.TableName, ex.Message));
                }

                return rez;
            }

            /// <summary>
            /// Обновление данных элемента конфигурации.
            /// </summary>
            /// <param name="IndexId">Индекс элемента который нужно обновить</param>
            /// <param name="updTbl">Пользователь у которого нужно изменить данные</param>
            /// <param name="HashExeption">C отображением исключений</param>
            /// <returns>Результат операции (Успех или нет)</returns>
            public bool Update(int IndexId, TableBase updTbl, bool HashExeption)
            {
                bool rez = false;
                try
                {
                    lock (this.TblL)
                    {

                        if (IndexId >= this.TblL.Count)
                        {
                            if (HashExeption) throw new ApplicationException(string.Format("Не удалось обновить данные элемента в списке {0}. Элемента с таким индексом {1} не существует.", updTbl.TableName, updTbl.ToString()));
                        }
                        else
                        {
                            updTbl.Index = IndexId;
                            this.TblL[IndexId] = updTbl;

                            rez = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (HashExeption) throw new ApplicationException(string.Format("Не удалось обновить данные элемента в списке {0}. Произошла ошибка: {1}", updTbl.TableName, ex.Message));
                }

                return rez;
            }

            /// <summary>
            /// Получение компонента по его ID
            /// </summary>
            /// <param name="i">Введите идентификатор</param>
            /// <returns></returns>
            public TableBase getTableComponent(int i)
            {
                try
                {
                    TableBase rez = null;
                    lock (TblL)
                    {
                        rez = this.TblL[i];
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
            public TableBase getTableComponent(string s)
            {
                try
                {
                    TableBase rez = null;
                    lock (TblL)
                    {
                        foreach (TableBase item in this.TblL)
                        {
                            if (item.TableName == s)
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
                lock (TblL)
                {
                    rez = this.TblL.GetEnumerator();
                }
                return rez;
            }
        }

    }
}
