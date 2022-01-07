using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;       // https://razilov-code.ru/2017/12/13/microsoft-office-interop-excel/  https://www.nookery.ru/c-work-c-excel/   http://www.wladm.narod.ru/C_Sharp/comexcel.html  https://docs.microsoft.com/ru-ru/dotnet/csharp/programming-guide/interop/how-to-access-office-onterop-objects   https://stackoverflow.com/questions/53735258/create-table-in-excel-with-datatable-and-interop
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
                    this.DefaultPathSource = DefaultPathSource ?? FarmExcel.DefaultPathSource;
                    this.DefaultPathTarget = DefaultPathTarget ?? FarmExcel.DefaultPathTarget;
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
                    Object templatePathObj = (Tsk.Source.IndexOf(@"\") > 0 ? Tsk.Source : string.Format(@"{0}\{1}", this.DefaultPathSource, Tsk.Source));
                    Object pathToSaveObj = (Tsk.Target.IndexOf(@"\") > 0 ? Tsk.Target : string.Format(@"{0}\{1}", this.DefaultPathTarget, Tsk.Target));

                    // Проверка путей
                    if (templatePathObj == null || string.IsNullOrWhiteSpace(templatePathObj.ToString())) throw new ApplicationException(string.Format("Не указан файл шаблона"));
                    if (!File.Exists(templatePathObj.ToString())) throw new ApplicationException(string.Format("Шаблон не найден по пути: ({0})", templatePathObj.ToString()));
                    if (pathToSaveObj == null || string.IsNullOrWhiteSpace(pathToSaveObj.ToString())) throw new ApplicationException(string.Format("Не указан файл relf куда сохранить результат."));
                    string DirTmp = Path.GetDirectoryName(pathToSaveObj.ToString());
                    if (!Directory.Exists(DirTmp)) throw new ApplicationException(string.Format("Целевой директории в которой должен лежать файл не существует: ({0})", templatePathObj));
                    if (!TmpReplaseFileTarget && File.Exists(pathToSaveObj.ToString())) throw new ApplicationException(string.Format("В Целевой папке уже существует файл с таким именем: ({0})", pathToSaveObj.ToString()));

                    // открываем приложение Excel
                    Excel.Application exelApp = new Excel.Application();

                    //Отключить отображение окон с сообщениями
                    exelApp.DisplayAlerts = false;

                    object NullValue = System.Reflection.Missing.Value;
                    //Excel.Workbook document = exelApp.Workbooks.Open(templatePathObj.ToString(), 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //Excel.Workbook document = exelApp.Workbooks.Open(templatePathObj.ToString();//, "", true, false, 0, true, false, false);
                    //Excel.Workbook document = exelApp.Workbooks.Open(templatePathObj.ToString(), "", true, false, 0, true, false, false);
                    Excel.Workbook document = exelApp.Workbooks.Open(templatePathObj.ToString(), NullValue, NullValue, NullValue, NullValue, NullValue, NullValue, NullValue, NullValue, NullValue, NullValue, NullValue, NullValue, NullValue, NullValue);


                    try
                    {
                        int SheetCount = exelApp.Worksheets.Count;
                        List<dynamic[]> SheetCountL = new List<dynamic[]>();
                        for (int i = 0; i < SheetCount; i++)
                        {
                            //Получаем лист который выбрал пользователь
                            Excel.Worksheet sheettmp = (Excel.Worksheet)exelApp.Worksheets.get_Item(i+1);

                            // Сохраняем текущую ширину
                            dynamic[] colWith = new dynamic[100];
                            for (int ic = 0; ic < 100; ic++)
                            {
                                Excel.Range rtmp = sheettmp.get_Range(string.Format("A{0}", ic + 1), string.Format("A{0}", ic + 1));
                                colWith[ic] = rtmp.ColumnWidth;
                            }
                            SheetCountL.Add(colWith);
                        }

                        // Пробегаем по всем таблицам
                        foreach (Table item in Tsk.TblL)
                        {
                            //Получаем лист который выбрал пользователь
                            Excel.Worksheet sheet = (Excel.Worksheet)exelApp.Worksheets.get_Item(int.Parse(item.TableName.Split('|')[0]));

                            //Получаем ячейку самого левого угла в таблице
                            Excel.Range range = sheet.get_Range(item.TableName.Split('|')[1], item.TableName.Split('|')[1]);


                            // Инициируем объект статистики
                            base.SetInitTableInExcelAffected(Tsk, new RezultTaskAffectetdRow(item));

                            // range[1, 1]  Этот адрес относительный. Получается можно теперь в него вставлять данные
                            for (int iRow = 0; iRow < item.TableValue.Rows.Count; iRow++)
                            {
                                for (int iCol = 0; iCol < item.TableValue.Columns.Count; iCol++)
                                {
                                    range[iRow+1, iCol+1] = item.TableValue.Rows[iRow][iCol].ToString();
                                }

                                // Обновляем статистику по таблице
                                base.SetTableInExcelAffected(Tsk, iRow + 1);
                            }
                        }

                        // Восстанавливаем начальную ширину
                        for (int i = 0; i < SheetCount; i++)
                        {
                            dynamic[] colWith = SheetCountL[i];

                            //Получаем лист который выбрал пользователь
                            Excel.Worksheet sheettmp = (Excel.Worksheet)exelApp.Worksheets.get_Item(i + 1);
                                                       

                            for (int ic = 0; ic < 100; ic++)
                            {
                                try
                                {
                                    Excel.Range rtmp = sheettmp.get_Range(string.Format("A{0}", ic + 1), string.Format("A{0}", ic + 1));
                                    rtmp.ColumnWidth = colWith[ic];
                                }
                                catch (Exception) { }
                            }
                        }

                        


                        // выставляем флаг что задание завершено успешно
                        base.SetStatusTaskExcel(Tsk, EnStatusTask.Refresh);

                        // Обновляем все источники данных
                        document.RefreshAll();

                        //System.Threading.Thread.Sleep(8000);

                        // Сохраняем документ
                        exelApp.Application.ActiveWorkbook.SaveAs(pathToSaveObj.ToString(), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                       
                        // выставляем флаг что задание завершено успешно
                        base.SetStatusTaskExcel(Tsk, EnStatusTask.Save);

                        // Делаем видимыми все документы в этом приложении
                        //application.Visible = true;

                        // выставляем флаг что задание завершено успешно
                        base.SetStatusTaskExcel(Tsk, EnStatusTask.Success);
                    }
                    catch (Exception ex)
                    {
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
                            if (exelApp != null)
                            {
                                exelApp.Quit();
                                exelApp = null;
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
                base.SetStatusTaskExcel(Tsk, EnStatusTask.ERROR);

                throw new ApplicationException(string.Format("{0}.StartCreateReport   Упали с ошибкой: ({1})", obj.GetType().Name, ex.Message));
            }
        }



    }
}
