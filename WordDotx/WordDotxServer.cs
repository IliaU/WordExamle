using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace WordDotx
{
    /// <summary>
    /// Класс для создания сервера который будет обрабатывать запросы
    /// </summary>
    public class WordDotxServer
    {
        /// <summary>
        /// Внутренний объект нашего Сервера
        /// </summary>
        private WordDotxServer obj;

        /// <summary>
        /// Папка по умолчанию для нашего файла с источником шаблонов
        /// </summary>
        public string DefaultPathSource;

        /// <summary>
        /// Папка по умолчанию для нашего файла в который положим результат
        /// </summary>
        public string DefaultPathTarget;
        
        /// <summary>
        /// Конструктор для создания сервера
        /// </summary>
        /// <param name="DefaultPathSource">Папка по умолчанию для нашего файла с источником шаблонов</param>
        /// <param name="DefaultPathTarget">Папка по умолчанию для нашего файла в который положим результат</param>
        public WordDotxServer(string DefaultPathSource, string DefaultPathTarget)
        {
            try
            {
                if (obj==null)
                {
                    this.DefaultPathSource = DefaultPathSource;
                    this.DefaultPathTarget = DefaultPathTarget;
                    obj = this;
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}   Упали с ошибкой в конструкторе: ({1})", obj.GetType().Name, ex.Message));
            }
        }
        //
        /// <summary>
        /// Конструктор для создания сервера
        /// </summary>
        /// <param name="DefPathSorsAndTarget">Если путь один и для входящих файлов и исходящих</param>
        public WordDotxServer(string DefPathSorsAndTarget) : this(DefPathSorsAndTarget, DefPathSorsAndTarget)
        {
            try
            {

            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}   Упали с ошибкой в конструкторе: ({1})", obj.GetType().Name, ex.Message));
            }
        }
        //
        /// <summary>
        /// Конструктор для создания сервера
        /// </summary>
        public WordDotxServer() : this(Environment.CurrentDirectory, Environment.CurrentDirectory)
        {
            try
            {

            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}   Упали с ошибкой в конструкторе: ({1})", obj.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Процесс создания отчёта с подменой в шаблоне необходимых элементов на наши закладки и таблицы
        /// </summary>
        /// <param name="Source">Путь к файлу шаблона или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса</param>
        /// <param name="Target">Путь к файлу отчёта который создать по окончании работы или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса</param>
        /// <param name="BkmrkL">Список закладок которые мы будем использовать</param>
        /// <param name="TblL">Список таблиц который будем использовать</param>
        /// <param name="ReplaseFileTarget">Замена в папке назначения файла если уже ст таким именем файл существует</param>
        public void StartCreateReport(string Source, string Target, BookmarkList BkmrkL, TableList TblL, bool ReplaseFileTarget)
        {
            try
            {
                // Процесс должен идти в один поток скорее всего работать в несколько может не получиться
                lock (obj)
                {

                    // создаём переменные
                    Object missingObj = System.Reflection.Missing.Value;
                    Object trueObj = true;
                    Object falseObj = false;
                    Object templatePathObj = (Source.IndexOf(@"\") > 0 ? Source : string.Format(@"{0}\{1}", this.DefaultPathSource, Source));
                    Object pathToSaveObj = (Target.IndexOf(@"\") > 0 ? Source : string.Format(@"{0}\{1}", this.DefaultPathTarget, Target));

                    // Проверка путей
                    if (templatePathObj==null || string.IsNullOrWhiteSpace(templatePathObj.ToString())) throw new ApplicationException(string.Format("Не указан файл шаблона"));
                    if (!File.Exists(templatePathObj.ToString())) throw new ApplicationException(string.Format("Шаблон не найден по пути: ({0})", templatePathObj.ToString()));
                    if (pathToSaveObj==null || string.IsNullOrWhiteSpace(pathToSaveObj.ToString())) throw new ApplicationException(string.Format("Не указан файл relf куда сохранить результат."));
                    string DirTmp = Path.GetDirectoryName(pathToSaveObj.ToString());
                    if (!Directory.Exists(DirTmp)) throw new ApplicationException(string.Format("Целевой директории в которой должен лежать файл не существует: ({0})", templatePathObj));
                    if (!ReplaseFileTarget && File.Exists(pathToSaveObj.ToString())) throw new ApplicationException(string.Format("В Целевой папке уже существует файл с таким именем: ({0})", pathToSaveObj.ToString()));

                    // открываем приложение ворда
                    Word._Application application = new Word.Application();
                    Word._Document document = null;

                    try
                    {
                        // Добавляем в приложение сам документ  обрати внимание что может быть несколько документов добавлено а потом в самом концеможно из сделать видимыми
                        document = application.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
                        //document = application.Documents.Add(ref templatePathObj1, ref missingObj, ref missingObj, ref missingObj);

                        // Находим все закладки и меняем в них значения
                        foreach (Bookmark item in BkmrkL)
                        {
                            document.Bookmarks[item.BookmarkName].Range.Text = item.BookmarkValue;
                        }

                        // Которые нам передали
                        foreach (Table item in TblL)
                        {
                            // Пробегаем по таблицам в корне
                            foreach (Word.Table itemT in document.Tables)
                            {
                                // Вызываем процесс обработки таблиц
                                ProcessTable(item, itemT);
                            }
                        }

                        // Сохраняем но как вордовский докумен
                        document.SaveAs(ref pathToSaveObj, Word.WdSaveFormat.wdFormatDocument);

                        // Делаем видимыми все документы в этом приложении
                        //application.Visible = true;

                    }
                    catch (Exception ex)
                    {
                        if (document != null)
                        {
                            document.Close(ref falseObj, ref missingObj, ref missingObj);
                            application.Quit(ref missingObj, ref missingObj, ref missingObj);
                        }
                        //throw ex;
                    }
                    finally
                    {
                        try
                        {
                            if (document != null)
                            {
                                document.Close();
                                document = null;
                            }
                            if (application != null)
                            {
                                application.Quit();
                                application = null;
                            }
                        }
                        catch (Exception) { }
                    }
                }

            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.StartCreateReport   Упали с ошибкой: ({1})", obj.GetType().Name, ex.Message));
            }
        }

        /// <summary>
        /// Процесс создания отчёта с подменой в шаблоне необходимых элементов на наши закладки и таблицы
        /// </summary>
        /// <param name="Source">Путь к файлу шаблона или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса</param>
        /// <param name="Target">Путь к файлу отчёта который создать по окончании работы или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса</param>
        /// <param name="Bkmrk">Закладока которые мы будем использовать</param>
        /// <param name="Tbl">Таблица который будем использовать</param>
        /// <param name="ReplaseFileTarget">Замена в папке назначения файла если уже ст таким именем файл существует</param>
        public void StartCreateReport(string Source, string Target, Bookmark Bkmrk, Table Tbl, bool ReplaseFileTarget)
        {
            try
            {
                BookmarkList BkmrkL = new BookmarkList();
                BkmrkL.Add(Bkmrk, true);

                TableList TblL = new TableList();
                TblL.Add(Tbl, true);

                StartCreateReport(Source, Target, BkmrkL, TblL, ReplaseFileTarget);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.StartCreateReport   Упали с ошибкой: ({1})", obj.GetType().Name, ex.Message));
            }
        }
        //
        /// <summary>
        /// Процесс создания отчёта с подменой в шаблоне необходимых элементов на наши закладки и таблицы
        /// </summary>
        /// <param name="Source">Путь к файлу шаблона или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса</param>
        /// <param name="Target">Путь к файлу отчёта который создать по окончании работы или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса</param>
        /// <param name="BkmrkL">Закладока которые мы будем использовать</param>
        /// <param name="Tbl">Таблица который будем использовать</param>
        /// <param name="ReplaseFileTarget">Замена в папке назначения файла если уже ст таким именем файл существует</param>
        public void StartCreateReport(string Source, string Target, BookmarkList BkmrkL, Table Tbl, bool ReplaseFileTarget)
        {
            try
            {
                TableList TblL = new TableList();
                TblL.Add(Tbl, true);

                StartCreateReport(Source, Target, BkmrkL, TblL, ReplaseFileTarget);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.StartCreateReport   Упали с ошибкой: ({1})", obj.GetType().Name, ex.Message));
            }
        }
        //
        /// <summary>
        /// Процесс создания отчёта с подменой в шаблоне необходимых элементов на наши закладки и таблицы
        /// </summary>
        /// <param name="Source">Путь к файлу шаблона или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса</param>
        /// <param name="Target">Путь к файлу отчёта который создать по окончании работы или имя файла тогда папка будет использоваться заданная по умолчанию при инициализации класса</param>
        /// <param name="Bkmrk">Закладока которые мы будем использовать</param>
        /// <param name="TblL">Таблица который будем использовать</param>
        /// <param name="ReplaseFileTarget">Замена в папке назначения файла если уже ст таким именем файл существует</param>
        public void StartCreateReport(string Source, string Target, Bookmark Bkmrk, TableList TblL, bool ReplaseFileTarget)
        {
            try
            {
                BookmarkList BkmrkL = new BookmarkList();
                BkmrkL.Add(Bkmrk, true);

                StartCreateReport(Source, Target, BkmrkL, TblL, ReplaseFileTarget);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.StartCreateReport   Упали с ошибкой: ({1})", obj.GetType().Name, ex.Message));
            }
        }


        /// <summary>
        /// Обработка конкретной таблицы оказывается там может быть вложенность таблиц
        /// </summary>
        /// <param name="Tab">Таблица с именем и индексом которую нужно найти и заменить</param>
        /// <param name="itemT">Таблица которую обрабатываем</param>
        private void ProcessTable(Table Tab, Word.Table itemT)
        {
            try
            {
                for (int i = 0; i < itemT.Rows.Count; i++)
                {
                    string FlagAddRow = null;

                    for (int ic = 0; ic < itemT.Rows[i + 1].Cells.Count; ic++)
                    {
                        // получаем содержимое
                        string tmpCell = itemT.Rows[i + 1].Cells[ic + 1].Range.Text;
                        if (itemT.Rows[i + 1].Cells[ic + 1].Tables.Count > 0)
                        {
                            // сначала перестраиваем внутренние таблицы
                            foreach (Word.Table itemTTinput in itemT.Rows[i + 1].Cells[ic + 1].Tables)
                            {
                                // обработка внутренней таблицы
                                ProcessTable(Tab, itemTTinput);
                            }

                            // Заново перечитаем переменную
                            tmpCell = itemT.Rows[i + 1].Cells[ic + 1].Range.Text;

                            // И вот теперь подмениваем с правильным содержимым
                            foreach (Word.Table itemTTinput in itemT.Rows[i + 1].Cells[ic + 1].Tables)
                            {
                                // Вот тут сложнее. Надо убрать все символы внутренней таблицы между внутри нашего текста для того чтобы мы не отреогировали и не вставили строку когда она относится ко внутренней таблице
                                string DelitStr = itemTTinput.Range.Text;
                                tmpCell = tmpCell.Replace(DelitStr, "");
                            }
                        }

                        // Поиск системных символов
                        if (tmpCell.IndexOf("\r\r") == 0) tmpCell = tmpCell.Substring(2);                                        // вначале ячейки встаёт системные символы их учитывать не надо
                        if (tmpCell.IndexOf("\r\a") == tmpCell.Length - 2) tmpCell = tmpCell.Substring(0, tmpCell.Length - 2);   // вконце ячейки встаёт системные символы их учитывать не надо
                        tmpCell = tmpCell.Replace("\r", "");  // режем системные символы они нам тут не к чему

                        //Если найден объект который потенциально может быть в нашем датасете
                        if (tmpCell.IndexOf("{@D") > -1)
                        {
                            FlagAddRow = tmpCell.Substring(tmpCell.IndexOf("{@D"), tmpCell.IndexOf(".", tmpCell.IndexOf("{@D")) - tmpCell.IndexOf("{@D"));

                            // тогда нужно выбрать какой подиток вывести
                            if (tmpCell.IndexOf(FlagAddRow + ".T") > -1)
                            {
                                // Вырезаем имя тотала
                                string TmpTotalColumn = tmpCell.Substring(tmpCell.IndexOf(FlagAddRow + ".T") + FlagAddRow.Length + 2, tmpCell.IndexOf("}", tmpCell.IndexOf("}")) - tmpCell.IndexOf(FlagAddRow + ".T") - +FlagAddRow.Length - 2);

                                //Если найден объект который потенциально может быть итогом в нашем датасете
                                if (TmpTotalColumn == "3")
                                {
                                    tmpCell = tmpCell.Replace(FlagAddRow + ".T" + TmpTotalColumn + "}", "Итог новый");
                                    itemT.Rows[i + 1].Cells[ic + 1].Range.Text = tmpCell;
                                }

                                // Если это тотал то не нужно разрезать на строки
                                FlagAddRow = null;
                            }
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(FlagAddRow))
                    {
                        // пробегаем по строкам которые нужно воткнуть
                        for (int io = 0; io < Tab.TableValue.Rows.Count; io++)
                        {
                            // не понял логики но тут вставит строку до той которую мы нашли
                            itemT.Rows.Add(itemT.Rows[i + io + 1]);

                            // пробегаем поячейкам походу так попадаем на нашу вставленную строку так как она вотнётся до той что копировали
                            // foreach (Word.Cell item in itemT.Rows[i + 1].Cells)
                            for (int ic = 0; ic < itemT.Rows[i + 1].Cells.Count; ic++)
                            {
                                // Походу тут отсчёт идёт не от нуля а от еденицы по крайней мере в колонках
                                string tmpCell = itemT.Rows[i + io + 2].Cells[ic + 1].Range.Text;
                                if (!string.IsNullOrWhiteSpace(tmpCell))
                                {
                                    if (tmpCell.IndexOf("\r\r") == 0) tmpCell = tmpCell.Substring(2);                                        // вначале ячейки встаёт системные символы их учитывать не надо
                                    if (tmpCell.IndexOf("\r\a") == tmpCell.Length - 2) tmpCell = tmpCell.Substring(0, tmpCell.Length - 2);   // вконце ячейки встаёт системные символы их учитывать не надо
                                    if (tmpCell.IndexOf("\r\a") == 0 && tmpCell.Length == 2) tmpCell = tmpCell.Replace("\r", "");              // В пустых колонках иногда бывате такая комбинация

                                    for (int ColI = 0; ColI < Tab.TableValue.Columns.Count; ColI++)
                                    {
                                        tmpCell = tmpCell.Replace(FlagAddRow + ".C"+ ColI + "}", Tab.TableValue.Rows[io][ColI].ToString());
                                    }

                                    itemT.Rows[i + io + 1].Cells[ic + 1].Range.Text = tmpCell;
                                }
                            }

                        }

                        // Удаляем нашу строку с шаблоном
                        itemT.Rows[i + 1 + Tab.TableValue.Rows.Count].Delete();

                        // Перепрыгиваем на следующую строку чтобы не обрабатыать повтороно вставленные строки
                        i = i + Tab.TableValue.Rows.Count - 1;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
