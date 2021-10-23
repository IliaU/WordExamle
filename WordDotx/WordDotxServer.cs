using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Runtime.InteropServices;

namespace WordDotx
{
    /// <summary>
    /// Класс для создания сервера который будет обрабатывать запросы
    /// </summary>
    public class WordDotxServer: Lib.TaskWordBase.RezultTaskBase.WordDotxServerBase
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
        /// Поведение по умолчанию нужно заменить файл или нет
        /// </summary>
        public bool DefReplaseFileTarget;

        /// <summary>
        /// Конструктор для создания сервера
        /// </summary>
        /// <param name="DefaultPathSource">Папка по умолчанию для нашего файла с источником шаблонов</param>
        /// <param name="DefaultPathTarget">Папка по умолчанию для нашего файла в который положим результат</param>
        /// <param name="DefReplaseFileTarget">Папка по умолчанию для нашего файла в который положим результат</param>
        public WordDotxServer(string DefaultPathSource, string DefaultPathTarget, bool DefReplaseFileTarget)
        {
            try
            {
                if (obj == null)
                {
                    // Если при создании сервера не указана папка по умолчанию то береём из Фарма
                    this.DefaultPathSource = DefaultPathSource ?? FarmWordDotx.DefaultPathSource;
                    this.DefaultPathTarget = DefaultPathTarget ?? FarmWordDotx.DefaultPathTarget;
                    this.DefReplaseFileTarget = DefReplaseFileTarget;
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
        /// <param name="DefaultPathSource">Папка по умолчанию для нашего файла с источником шаблонов</param>
        /// <param name="DefaultPathTarget">Папка по умолчанию для нашего файла в который положим результат</param>
        public WordDotxServer(string DefaultPathSource, string DefaultPathTarget) : this(DefaultPathSource, DefaultPathTarget, true)
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
        /// <param name="DefPathSorsAndTarget">Если путь один и для входящих файлов и исходящих</param>
        public WordDotxServer(string DefPathSorsAndTarget) : this(DefPathSorsAndTarget, DefPathSorsAndTarget, true)
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
        public WordDotxServer() : this(Environment.CurrentDirectory, Environment.CurrentDirectory, true)
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
        /// <param name="Tsk">Задание которое нужно выполнить</param>
        public void StartCreateReport(TaskWord Tsk)
        {
            try
            {
                if (Tsk.StatusTask == EnStatusTask.Pending) throw new ApplicationException("Данное задание уже находится в очереди асинхронной обработки и должно запускаться через Пул обработки асинхронных запросов."); 

                // выставляем флаг что задание начинает свою работу
                base.SetStatusTaskWord(Tsk, EnStatusTask.Running);

                bool TmpReplaseFileTarget = this.DefReplaseFileTarget;
                if (Tsk.ReplaseFileTarget != null) TmpReplaseFileTarget = (bool)Tsk.ReplaseFileTarget;

                // Процесс должен идти в один поток скорее всего работать в несколько может не получиться
                lock (obj)
                {
                    // создаём переменные
                    Object missingObj = System.Reflection.Missing.Value;
                    Object trueObj = true;
                    Object falseObj = false;
                    Object templatePathObj = (Tsk.Source.IndexOf(@"\") > 0 ? Tsk.Source : string.Format(@"{0}\{1}", this.DefaultPathSource, Tsk.Source));
                    Object pathToSaveObj = (Tsk.Target.IndexOf(@"\") > 0 ? Tsk.Target : string.Format(@"{0}\{1}", this.DefaultPathTarget, Tsk.Target));

                    // Проверка путей
                    if (templatePathObj == null || string.IsNullOrWhiteSpace(templatePathObj.ToString())) throw new ApplicationException(string.Format("Не указан файл шаблона"));
                    if (!File.Exists(templatePathObj.ToString())) throw new ApplicationException(string.Format("Шаблон не найден по пути: ({0})", templatePathObj.ToString()));
                    if (pathToSaveObj == null || string.IsNullOrWhiteSpace(pathToSaveObj.ToString())) throw new ApplicationException(string.Format("Не указан файл relf куда сохранить результат."));
                    string DirTmp = Path.GetDirectoryName(pathToSaveObj.ToString());
                    if (!Directory.Exists(DirTmp)) throw new ApplicationException(string.Format("Целевой директории в которой должен лежать файл не существует: ({0})", templatePathObj));
                    if (!TmpReplaseFileTarget && File.Exists(pathToSaveObj.ToString())) throw new ApplicationException(string.Format("В Целевой папке уже существует файл с таким именем: ({0})", pathToSaveObj.ToString()));

                    // открываем приложение ворда
                    Word._Application application = new Word.Application();
                    Word._Document document = null;

                    try
                    {
                        // Добавляем в приложение сам документ  обрати внимание что может быть несколько документов добавлено а потом в самом концеможно из сделать видимыми
                        document = application.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
                        //document = application.Documents.Add(ref templatePathObj1, ref missingObj, ref missingObj, ref missingObj);

                        // Находим все закладки и меняем в них значения
                        foreach (Bookmark item in Tsk.BkmrkL)
                        {
                            try
                            {
                                document.Bookmarks[item.BookmarkName].Range.Text = item.BookmarkValue;
                            }
                            catch (Exception ex)
                            {
                                // throw;Тут надо прикрутить событие варнинг типо не используется в этом файле такая закладка не найдена
                            }
                        }

                        // Которые нам передали
                        foreach (Table item in Tsk.TblL)
                        {
                            // Пробегаем по таблицам в корне
                            foreach (Word.Table itemT in document.Tables)
                            {
                                // Вызываем процесс обработки таблиц
                                ProcessTable(Tsk, item, itemT, document);
                            }
                        }

                        // выставляем флаг что задание завершено успешно
                        base.SetStatusTaskWord(Tsk, EnStatusTask.Save);

                        // Сохраняем но как вордовский докумен
                        document.SaveAs(ref pathToSaveObj, Word.WdSaveFormat.wdFormatDocument);

                        // Делаем видимыми все документы в этом приложении
                        //application.Visible = true;

                        // выставляем флаг что задание завершено успешно
                        base.SetStatusTaskWord(Tsk, EnStatusTask.Success);
                    }
                    catch (Exception ex)
                    {
                        if (document != null)
                        {
                            document.Close(ref falseObj, ref missingObj, ref missingObj);
                            application.Quit(ref missingObj, ref missingObj, ref missingObj);
                        }
                        throw ex;
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
                // выставляем флаг что задание завершено c ошибкой
                base.SetStatusMessage(Tsk, ex.Message);
                base.SetStatusTaskWord(Tsk, EnStatusTask.ERROR);

                throw new ApplicationException(string.Format("{0}.StartCreateReport   Упали с ошибкой: ({1})", obj.GetType().Name, ex.Message));
            }
        }


        /// <summary>
        /// Обработка конкретной таблицы оказывается там может быть вложенность таблиц
        /// </summary>
        /// <param name="Tab">Таблица с именем и индексом которую нужно найти и заменить</param>
        /// <param name="itemT">Таблица которую обрабатываем</param>
        /// <param name="Doc">Документ</param>
        private void ProcessTable(TaskWord Tsk, Table Tab, Word.Table itemT, Word._Document Doc)
        {
            try
            {
                // Инициируем объект статистики
                base.SetInitTableInWordAffected(Tsk, new RezultTaskAffectetdRow(Tab));
                int FlagPart = 100; // какой промежуток строк через который нужно обновить стату для того чтобы не часто срабатывали события
                int FlatPartTmp = FlagPart; // Текущее значение счётчика

                for (int i = 0; i < itemT.Rows.Count; i++)
                {
                    // Правим статистику у последней таблицы
                    if (FlatPartTmp > 0) FlatPartTmp--;
                    else
                    {
                        base.SetTableInWordAffected(Tsk, itemT.Rows.Count);
                        FlatPartTmp = FlagPart;
                    }

                    string FlagAddRow = null;

                    for (int ic = 0; ic < itemT.Columns.Count; ic++)
                    {
                        try
                        {
                            // получаем содержимое
                            string tmpCell = itemT.Cell(i + 1, ic + 1).Range.Text;
                            if (itemT.Cell(i + 1, ic + 1).Tables.Count > 0)
                            {
                                // сначала перестраиваем внутренние таблицы
                                foreach (Word.Table itemTTinput in itemT.Cell(i + 1, ic + 1).Tables)
                                {
                                    // обработка внутренней таблицы
                                    ProcessTable(Tsk, Tab, itemTTinput, Doc);
                                }

                                // Заново перечитаем переменную
                                tmpCell = itemT.Cell(i + 1, ic + 1).Range.Text;

                                // И вот теперь подмениваем с правильным содержимым
                                foreach (Word.Table itemTTinput in itemT.Cell(i + 1, ic + 1).Tables)
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
                                // Получаем флаг
                                string TmpFlagAddRow = tmpCell.Substring(tmpCell.IndexOf("{@D"), tmpCell.IndexOf(".", tmpCell.IndexOf("{@D")) - tmpCell.IndexOf("{@D"));

                                // Получаем имя таблицы и проверяем что она соответствует той что мы сейчас смотрим
                                string TmpTableName = TmpFlagAddRow.Substring(tmpCell.IndexOf("D") + 1, TmpFlagAddRow.Length - tmpCell.IndexOf("D") - 1);
                                int TmpTableIndex = -1;
                                try { TmpTableIndex = int.Parse(TmpTableName); }
                                catch (Exception) { }
                                if (TmpTableIndex == -1)     //  Если в имени датасета указан не индекс а имя таблицы
                                {   // Имя таблицы совподает занчит надо взвести наш флаг
                                    if (TmpTableName == Tab.TableName) FlagAddRow = TmpFlagAddRow;
                                }
                                else                         // В имени датасета указан индекс а не имя таблицы
                                {   // Индекс таблицы совподает значит надо взвести наш флаг
                                    if (TmpTableIndex == Tab.Index) FlagAddRow = TmpFlagAddRow;
                                }

                                // Если обнаружена именно та таблица которыю мы обрабатываем
                                if (!string.IsNullOrWhiteSpace(FlagAddRow))
                                {
                                    // тогда нужно выбрать какой подиток вывести
                                    if (tmpCell.IndexOf(FlagAddRow + ".T") > -1)
                                    {
                                        // Вырезаем имя тотала
                                        bool TmpFlatCurTotal = false;
                                        string TmpTotalColumn = tmpCell.Substring(tmpCell.IndexOf(FlagAddRow + ".T") + FlagAddRow.Length + 2, tmpCell.IndexOf("}", tmpCell.IndexOf("}")) - tmpCell.IndexOf(FlagAddRow + ".T") - +FlagAddRow.Length - 2);
                                        int TmpTotalIndex = -1;
                                        try { TmpTotalIndex = int.Parse(TmpTotalColumn); }
                                        catch (Exception) { }
                                        // Пробегаем по всем тоталам
                                        foreach (Total itemTtl in Tab.TtlList)
                                        {
                                            if (TmpTotalIndex == -1)     //  Если в имени датасета указан не индекс а имя таблицы
                                            {   // Имя таблицы совподает занчит надо взвести наш флаг
                                                if (TmpTotalColumn == itemTtl.TotalName) TmpFlatCurTotal = true;
                                            }
                                            else                         // В имени датасета указан индекс а не имя таблицы
                                            {   // Индекс таблицы совподает значит надо взвести наш флаг
                                                if (TmpTotalIndex == itemTtl.Index) TmpFlatCurTotal = true;
                                            }

                                            //Если найден объект который является тоталом
                                            if (TmpFlatCurTotal)
                                            {
                                                tmpCell = tmpCell.Replace(FlagAddRow + ".T" + TmpTotalColumn + "}", itemTtl.TotalValue);
                                                itemT.Cell(i + 1, ic + 1).Range.Text = tmpCell;

                                                // Если в этой ячейке нашли тотал то нет смысла перебирать и искать другие тоталы которые существуют в этой табичке
                                                break;
                                            }
                                        }

                                        // Если это тотал то не нужно разрезать на строки в любом случае
                                        FlagAddRow = null;
                                    }
                                }
                            }
                        }
                        catch (COMException ex)
                        {
                            switch (ex.ErrorCode)
                            {
                                case -2146822347:   // Запрашиваемый номер семейства не существует.
                                    break;
                                default:
                                    throw ex;
                            }
                        }
                        catch (Exception ex) { throw ex; }
                    }

                    if (!string.IsNullOrWhiteSpace(FlagAddRow))
                    {
                        // пробегаем по строкам которые нужно воткнуть
                        for (int io = 0; io < Tab.TableValue.Rows.Count; io++)
                        {
                            try
                            {
                                // не понял логики но тут вставит строку до той которую мы нашли
                                //itemT.Rows.Add(itemT.Rows[i + io + 1]);
                                //itemT.Rows.Add(itemT.Cell(i + io + 1, 1).Row);
                                // При этих вариантах падает с ошибкой Отсутствует доступ к отдельным строкам, поскольку таблица имеет ячейки, объединенные по вертикали. зато можно найти любой диапазон представить в виде строки и его добавить
                                Doc.Range(itemT.Cell(i + io + 1, 1).Range.Start, itemT.Cell(i + io + 1, itemT.Columns.Count).Range.End).Rows.Add(Doc.Range(itemT.Cell(i + io + 1, 1).Range.Start, itemT.Cell(i + io + 1, itemT.Columns.Count).Range.End).Rows);

                                // пробегаем поячейкам походу так попадаем на нашу вставленную строку так как она вотнётся до той что копировали
                                // foreach (Word.Cell item in itemT.Rows[i + 1].Cells)
                                for (int ic = 0; ic < itemT.Columns.Count; ic++)
                                {
                                    // Походу тут отсчёт идёт не от нуля а от еденицы по крайней мере в колонках
                                    string tmpCell = itemT.Cell(i + io + 2, ic + 1).Range.Text;
                                    if (!string.IsNullOrWhiteSpace(tmpCell))
                                    {
                                        if (tmpCell.IndexOf("\r\r") == 0) tmpCell = tmpCell.Substring(2);                                        // вначале ячейки встаёт системные символы их учитывать не надо
                                        if (tmpCell.IndexOf("\r\a") == tmpCell.Length - 2) tmpCell = tmpCell.Substring(0, tmpCell.Length - 2);   // вконце ячейки встаёт системные символы их учитывать не надо
                                        if (tmpCell.IndexOf("\r\a") == 0 && tmpCell.Length == 2) tmpCell = tmpCell.Replace("\r", "");              // В пустых колонках иногда бывате такая комбинация

                                        for (int ColI = 0; ColI < Tab.TableValue.Columns.Count; ColI++)
                                        {
                                            // запоминаем начальное значение ячейки
                                            string tmpCellStart = tmpCell;

                                            // Подмена по индексу колонки
                                            tmpCell = tmpCell.Replace(string.Format("{{@D{0}.C{1}}}", Tab.TableName, ColI), Tab.TableValue.Rows[io][ColI].ToString());

                                            // подмена по имени колонки
                                            tmpCell = tmpCell.Replace(string.Format("{{@D{0}.C{1}}}", Tab.TableName, Tab.TableValue.Columns[ColI].ColumnName), Tab.TableValue.Rows[io][ColI].ToString());

                                            // Подмена по индексу колонки
                                            tmpCell = tmpCell.Replace(string.Format("{{@D{0}.C{1}}}", Tab.Index, ColI), Tab.TableValue.Rows[io][ColI].ToString());

                                            // подмена по имени колонки
                                            tmpCell = tmpCell.Replace(string.Format("{{@D{0}.C{1}}}", Tab.Index, Tab.TableValue.Columns[ColI].ColumnName), Tab.TableValue.Rows[io][ColI].ToString());

                                            // если ячейка изменилась значит найдено необходимое значенние и другие значения искать не нужно
                                            if (tmpCell != tmpCellStart) break;
                                        }

                                        itemT.Cell(i + io + 1, ic + 1).Range.Text = tmpCell;
                                    }
                                }

                                // предпологаю что если есть группировка, то надо где то в этом месте делать объединение с предыдужей ячейкой

                            }
                            catch (COMException ex)
                            {
                                switch (ex.ErrorCode)
                                {
                                    case -2146822347:   // Запрашиваемый номер семейства не существует.
                                        break;
                                    default:
                                        throw ex;
                                }
                            }
                            catch (Exception ex) { throw ex; }

                        }

                        // Удаляем нашу строку с шаблоном
                        //itemT.Rows[i + 1 + Tab.TableValue.Rows.Count].Delete();
                        // При этих вариантах падает с ошибкой Отсутствует доступ к отдельным строкам, поскольку таблица имеет ячейки, объединенные по вертикали. Зато можно найти любой диапазон представить в виде строки и его удалить
                        Doc.Range(itemT.Cell(i + 1 + Tab.TableValue.Rows.Count, 1).Range.Start, itemT.Cell(i + 1 + Tab.TableValue.Rows.Count, itemT.Columns.Count).Range.End).Rows.Delete();

                        // Перепрыгиваем на следующую строку чтобы не обрабатыать повтороно вставленные строки
                        i = i + Tab.TableValue.Rows.Count - 1;
                    }
                }

                // Правим статистику у последней таблицы и выставляем флаг что она готова
                base.SetEndTableInWordAffected(Tsk, itemT.Rows.Count);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
