using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;

namespace WordDotx
{
    /// <summary>
    /// Класс для создания сервера который будет обрабатывать запросы
    /// </summary>
    public class ExcelServer: Lib.TaskExcelBase.RezultTaskBase.ExcelServerBase
    {
        /// <summary>
        /// Внутренний объект нашего Сервера
        /// </summary>
        private ExcelServer obj;

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
        public ExcelServer(string DefaultPathSource, string DefaultPathTarget, bool DefReplaseFileTarget)
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
        public ExcelServer(string DefaultPathSource, string DefaultPathTarget) : this(DefaultPathSource, DefaultPathTarget, true)
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
        public ExcelServer(string DefPathSorsAndTarget) : this(DefPathSorsAndTarget, DefPathSorsAndTarget, true)
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
        public ExcelServer() : this(Environment.CurrentDirectory, Environment.CurrentDirectory, true)
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
        public void StartCreateReport(TaskExcel Tsk)
        {
            try
            {
                if (Tsk.StatusTask == EnStatusTask.Pending) throw new ApplicationException("Данное задание уже находится в очереди асинхронной обработки и должно запускаться через Пул обработки асинхронных запросов.");

                // выставляем флаг что задание начинает свою работу
                base.SetStatusTaskExcel(Tsk, EnStatusTask.Running);

                bool TmpReplaseFileTarget = this.DefReplaseFileTarget;
                if (Tsk.ReplaseFileTarget != null) TmpReplaseFileTarget = (bool)Tsk.ReplaseFileTarget;

                // Процесс должен идти в один поток скорее всего работать в несколько может не получиться
                lock (obj)
                {
                    // создаём переменные
                    //Object missingObj = System.Reflection.Missing.Value;
                    //Object trueObj = true;
                    //Object falseObj = false;
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
//                    Word._Application application = new Word.Application();
//                    Word._Document document = null;

                    try
                    {
                        // Добавляем в приложение сам документ  обрати внимание что может быть несколько документов добавлено а потом в самом концеможно из сделать видимыми
//                        document = application.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
                        //document = application.Documents.Add(ref templatePathObj1, ref missingObj, ref missingObj, ref missingObj);

                       
                        // Которые нам передали
                        foreach (Table item in Tsk.TblL)
                        {
                            // Пробегаем по таблицам в корне
              //              foreach (Word.Table itemT in document.Tables)
              //              {
                                // Вызываем процесс обработки таблиц
              //                  ProcessTable(Tsk, item, itemT, document);
              //              }
                        }

                        // выставляем флаг что задание завершено успешно
                        base.SetStatusTaskExcel(Tsk, EnStatusTask.Save);

                        // Сохраняем но как вордовский докумен
               //         document.SaveAs(ref pathToSaveObj, Excel.WdSaveFormat.wdFormatDocument);

                        // Делаем видимыми все документы в этом приложении
                        //application.Visible = true;

                        // выставляем флаг что задание завершено успешно
                        base.SetStatusTaskExcel(Tsk, EnStatusTask.Success);
                    }
                    catch (Exception ex)
                    {
      //                  if (document != null)
       //                 {
        //                    document.Close(ref falseObj, ref missingObj, ref missingObj);
        //                    application.Quit(ref missingObj, ref missingObj, ref missingObj);
        //                }
                        throw ex;
                    }
                    finally
                    {
                        try
                        {
     //                       if (document != null)
     //                       {
      //                          document.Close();
      //                          document = null;
      //                      }
      //                      if (application != null)
      //                      {
       //                         application.Quit();
        //                        application = null;
        //                    }
                        }
                        catch (Exception) { }
                    }
                }

            }
            catch (Exception ex)
            {
                // выставляем флаг что задание завершено c ошибкой
                base.SetStatusMessage(Tsk, ex.Message);
                base.SetStatusTaskExcel(Tsk, EnStatusTask.ERROR);

                throw new ApplicationException(string.Format("{0}.StartCreateReport   Упали с ошибкой: ({1})", obj.GetType().Name, ex.Message));
            }
        }


        /// <summary>
        /// Обработка конкретной таблицы оказывается там может быть вложенность таблиц
        /// </summary>
        /// <param name="Tab">Таблица с именем и индексом которую нужно найти и заменить</param>
        /// <param name="itemT">Таблица которую обрабатываем</param>
        /// <param name="Doc">Документ</param>
        private void ProcessTable(TaskExcel Tsk, Table Tab/*, Excel.Table itemT, Word._Document Doc*/)
        {
            try
            {
                // Инициируем объект статистики
                base.SetInitTableInExcelAffected(Tsk, new RezultTaskAffectetdRow(Tab));

                //for (int i = 0; i < itemT.Rows.Count; i++)
               // {
                    
               // }

                // Правим статистику у последней таблицы и выставляем флаг что она готова
                base.SetEndTableInExcelAffected(Tsk, 0 /* itemT.Rows.Count*/);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
