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
                            dynamic[] colWith = new dynamic[104];
                            colWith[0] = sheettmp.Columns["A:A"].ColumnWidth;
                            colWith[1] = sheettmp.Columns["B:B"].ColumnWidth;
                            colWith[2] = sheettmp.Columns["C:C"].ColumnWidth;
                            colWith[3] = sheettmp.Columns["D:D"].ColumnWidth;
                            colWith[4] = sheettmp.Columns["E:E"].ColumnWidth;
                            colWith[5] = sheettmp.Columns["F:F"].ColumnWidth;
                            colWith[6] = sheettmp.Columns["G:G"].ColumnWidth;
                            colWith[7] = sheettmp.Columns["H:H"].ColumnWidth;
                            colWith[8] = sheettmp.Columns["I:I"].ColumnWidth;
                            colWith[9] = sheettmp.Columns["G:G"].ColumnWidth;
                            colWith[10] = sheettmp.Columns["K:K"].ColumnWidth;
                            colWith[11] = sheettmp.Columns["L:L"].ColumnWidth;
                            colWith[12] = sheettmp.Columns["M:M"].ColumnWidth;
                            colWith[13] = sheettmp.Columns["N:N"].ColumnWidth;
                            colWith[14] = sheettmp.Columns["O:O"].ColumnWidth;
                            colWith[15] = sheettmp.Columns["P:P"].ColumnWidth;
                            colWith[16] = sheettmp.Columns["Q:Q"].ColumnWidth;
                            colWith[17] = sheettmp.Columns["R:R"].ColumnWidth;
                            colWith[18] = sheettmp.Columns["S:S"].ColumnWidth;
                            colWith[19] = sheettmp.Columns["T:T"].ColumnWidth;
                            colWith[20] = sheettmp.Columns["U:U"].ColumnWidth;
                            colWith[21] = sheettmp.Columns["V:V"].ColumnWidth;
                            colWith[22] = sheettmp.Columns["W:W"].ColumnWidth;
                            colWith[23] = sheettmp.Columns["X:X"].ColumnWidth;
                            colWith[24] = sheettmp.Columns["Y:Y"].ColumnWidth;
                            colWith[25] = sheettmp.Columns["Z:Z"].ColumnWidth;
                            colWith[26] = sheettmp.Columns["AA:AA"].ColumnWidth;
                            colWith[27] = sheettmp.Columns["AB:AB"].ColumnWidth;
                            colWith[28] = sheettmp.Columns["AC:AC"].ColumnWidth;
                            colWith[29] = sheettmp.Columns["AD:AD"].ColumnWidth;
                            colWith[30] = sheettmp.Columns["AE:AE"].ColumnWidth;
                            colWith[31] = sheettmp.Columns["AF:AF"].ColumnWidth;
                            colWith[32] = sheettmp.Columns["AG:AG"].ColumnWidth;
                            colWith[33] = sheettmp.Columns["AH:AH"].ColumnWidth;
                            colWith[34] = sheettmp.Columns["AI:AI"].ColumnWidth;
                            colWith[35] = sheettmp.Columns["AG:AG"].ColumnWidth;
                            colWith[36] = sheettmp.Columns["AK:AK"].ColumnWidth;
                            colWith[37] = sheettmp.Columns["AL:AL"].ColumnWidth;
                            colWith[38] = sheettmp.Columns["AM:AM"].ColumnWidth;
                            colWith[39] = sheettmp.Columns["AN:AN"].ColumnWidth;
                            colWith[40] = sheettmp.Columns["AO:AO"].ColumnWidth;
                            colWith[41] = sheettmp.Columns["AP:AP"].ColumnWidth;
                            colWith[42] = sheettmp.Columns["AQ:AQ"].ColumnWidth;
                            colWith[43] = sheettmp.Columns["AR:AR"].ColumnWidth;
                            colWith[44] = sheettmp.Columns["AS:AS"].ColumnWidth;
                            colWith[45] = sheettmp.Columns["AT:AT"].ColumnWidth;
                            colWith[46] = sheettmp.Columns["AU:AU"].ColumnWidth;
                            colWith[47] = sheettmp.Columns["AV:AV"].ColumnWidth;
                            colWith[48] = sheettmp.Columns["AW:AW"].ColumnWidth;
                            colWith[49] = sheettmp.Columns["AX:AX"].ColumnWidth;
                            colWith[50] = sheettmp.Columns["AY:AY"].ColumnWidth;
                            colWith[51] = sheettmp.Columns["AZ:AZ"].ColumnWidth;
                            colWith[52] = sheettmp.Columns["BA:BA"].ColumnWidth;
                            colWith[53] = sheettmp.Columns["BB:BB"].ColumnWidth;
                            colWith[54] = sheettmp.Columns["BC:BC"].ColumnWidth;
                            colWith[55] = sheettmp.Columns["BD:BD"].ColumnWidth;
                            colWith[56] = sheettmp.Columns["BE:BE"].ColumnWidth;
                            colWith[57] = sheettmp.Columns["BF:BF"].ColumnWidth;
                            colWith[58] = sheettmp.Columns["BG:BG"].ColumnWidth;
                            colWith[59] = sheettmp.Columns["BH:BH"].ColumnWidth;
                            colWith[60] = sheettmp.Columns["BI:BI"].ColumnWidth;
                            colWith[61] = sheettmp.Columns["BG:BG"].ColumnWidth;
                            colWith[62] = sheettmp.Columns["BK:BK"].ColumnWidth;
                            colWith[63] = sheettmp.Columns["BL:BL"].ColumnWidth;
                            colWith[64] = sheettmp.Columns["BM:BM"].ColumnWidth;
                            colWith[65] = sheettmp.Columns["BN:BN"].ColumnWidth;
                            colWith[66] = sheettmp.Columns["BO:BO"].ColumnWidth;
                            colWith[67] = sheettmp.Columns["BP:BP"].ColumnWidth;
                            colWith[68] = sheettmp.Columns["BQ:BQ"].ColumnWidth;
                            colWith[69] = sheettmp.Columns["BR:BR"].ColumnWidth;
                            colWith[70] = sheettmp.Columns["BS:BS"].ColumnWidth;
                            colWith[71] = sheettmp.Columns["BT:BT"].ColumnWidth;
                            colWith[72] = sheettmp.Columns["BU:BU"].ColumnWidth;
                            colWith[73] = sheettmp.Columns["BV:BV"].ColumnWidth;
                            colWith[74] = sheettmp.Columns["BW:BW"].ColumnWidth;
                            colWith[75] = sheettmp.Columns["BX:BX"].ColumnWidth;
                            colWith[76] = sheettmp.Columns["BY:BY"].ColumnWidth;
                            colWith[77] = sheettmp.Columns["BZ:BZ"].ColumnWidth;
                            colWith[78] = sheettmp.Columns["CA:CA"].ColumnWidth;
                            colWith[79] = sheettmp.Columns["CB:CB"].ColumnWidth;
                            colWith[80] = sheettmp.Columns["CC:CC"].ColumnWidth;
                            colWith[81] = sheettmp.Columns["CD:CD"].ColumnWidth;
                            colWith[82] = sheettmp.Columns["CE:CE"].ColumnWidth;
                            colWith[83] = sheettmp.Columns["CF:CF"].ColumnWidth;
                            colWith[84] = sheettmp.Columns["CG:CG"].ColumnWidth;
                            colWith[85] = sheettmp.Columns["CH:CH"].ColumnWidth;
                            colWith[86] = sheettmp.Columns["CI:CI"].ColumnWidth;
                            colWith[87] = sheettmp.Columns["CG:CG"].ColumnWidth;
                            colWith[88] = sheettmp.Columns["CK:CK"].ColumnWidth;
                            colWith[89] = sheettmp.Columns["CL:CL"].ColumnWidth;
                            colWith[90] = sheettmp.Columns["CM:CM"].ColumnWidth;
                            colWith[91] = sheettmp.Columns["CN:CN"].ColumnWidth;
                            colWith[92] = sheettmp.Columns["CO:CO"].ColumnWidth;
                            colWith[93] = sheettmp.Columns["CP:CP"].ColumnWidth;
                            colWith[94] = sheettmp.Columns["CQ:CQ"].ColumnWidth;
                            colWith[95] = sheettmp.Columns["CR:CR"].ColumnWidth;
                            colWith[96] = sheettmp.Columns["CS:CS"].ColumnWidth;
                            colWith[97] = sheettmp.Columns["CT:CT"].ColumnWidth;
                            colWith[98] = sheettmp.Columns["CU:CU"].ColumnWidth;
                            colWith[99] = sheettmp.Columns["CV:CV"].ColumnWidth;
                            colWith[100] = sheettmp.Columns["CW:CW"].ColumnWidth;
                            colWith[101] = sheettmp.Columns["CX:CX"].ColumnWidth;
                            colWith[102] = sheettmp.Columns["CY:CY"].ColumnWidth;
                            colWith[103] = sheettmp.Columns["CZ:CZ"].ColumnWidth;

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

                            try { sheettmp.Columns["A:A"].ColumnWidth = colWith[0]; } catch (Exception) { }
                            try { sheettmp.Columns["B:B"].ColumnWidth = colWith[1]; } catch (Exception) { }
                            try { sheettmp.Columns["C:C"].ColumnWidth = colWith[2]; } catch (Exception) { }
                            try { sheettmp.Columns["D:D"].ColumnWidth = colWith[3]; } catch (Exception) { }
                            try { sheettmp.Columns["E:E"].ColumnWidth = colWith[4]; } catch (Exception) { }
                            try { sheettmp.Columns["F:F"].ColumnWidth = colWith[5]; } catch (Exception) { }
                            try { sheettmp.Columns["G:G"].ColumnWidth = colWith[6]; } catch (Exception) { }
                            try { sheettmp.Columns["H:H"].ColumnWidth = colWith[7]; } catch (Exception) { }
                            try { sheettmp.Columns["I:I"].ColumnWidth = colWith[8]; } catch (Exception) { }
                            try { sheettmp.Columns["G:G"].ColumnWidth = colWith[9]; } catch (Exception) { }
                            try { sheettmp.Columns["K:K"].ColumnWidth = colWith[10]; } catch (Exception) { }
                            try { sheettmp.Columns["L:L"].ColumnWidth = colWith[11]; } catch (Exception) { }
                            try { sheettmp.Columns["M:M"].ColumnWidth = colWith[12]; } catch (Exception) { }
                            try { sheettmp.Columns["N:N"].ColumnWidth = colWith[13]; } catch (Exception) { }
                            try { sheettmp.Columns["O:O"].ColumnWidth = colWith[14]; } catch (Exception) { }
                            try { sheettmp.Columns["P:P"].ColumnWidth = colWith[15]; } catch (Exception) { }
                            try { sheettmp.Columns["Q:Q"].ColumnWidth = colWith[16]; } catch (Exception) { }
                            try { sheettmp.Columns["R:R"].ColumnWidth = colWith[17]; } catch (Exception) { }
                            try { sheettmp.Columns["S:S"].ColumnWidth = colWith[18]; } catch (Exception) { }
                            try { sheettmp.Columns["T:T"].ColumnWidth = colWith[19]; } catch (Exception) { }
                            try { sheettmp.Columns["U:U"].ColumnWidth = colWith[20]; } catch (Exception) { }
                            try { sheettmp.Columns["V:V"].ColumnWidth = colWith[21]; } catch (Exception) { }
                            try { sheettmp.Columns["W:W"].ColumnWidth = colWith[22]; } catch (Exception) { }
                            try { sheettmp.Columns["X:X"].ColumnWidth = colWith[23]; } catch (Exception) { }
                            try { sheettmp.Columns["Y:Y"].ColumnWidth = colWith[24]; } catch (Exception) { }
                            try { sheettmp.Columns["Z:Z"].ColumnWidth = colWith[25]; } catch (Exception) { }
                            try { sheettmp.Columns["AA:AA"].ColumnWidth = colWith[26]; } catch (Exception) { }
                            try { sheettmp.Columns["AB:AB"].ColumnWidth = colWith[27]; } catch (Exception) { }
                            try { sheettmp.Columns["AC:AC"].ColumnWidth = colWith[28]; } catch (Exception) { }
                            try { sheettmp.Columns["AD:AD"].ColumnWidth = colWith[29]; } catch (Exception) { }
                            try { sheettmp.Columns["AE:AE"].ColumnWidth = colWith[30]; } catch (Exception) { }
                            try { sheettmp.Columns["AF:AF"].ColumnWidth = colWith[31]; } catch (Exception) { }
                            try { sheettmp.Columns["AG:AG"].ColumnWidth = colWith[32]; } catch (Exception) { }
                            try { sheettmp.Columns["AH:AH"].ColumnWidth = colWith[33]; } catch (Exception) { }
                            try { sheettmp.Columns["AI:AI"].ColumnWidth = colWith[34]; } catch (Exception) { }
                            try { sheettmp.Columns["AG:AG"].ColumnWidth = colWith[35]; } catch (Exception) { }
                            try { sheettmp.Columns["AK:AK"].ColumnWidth = colWith[36]; } catch (Exception) { }
                            try { sheettmp.Columns["AL:AL"].ColumnWidth = colWith[37]; } catch (Exception) { }
                            try { sheettmp.Columns["AM:AM"].ColumnWidth = colWith[38]; } catch (Exception) { }
                            try { sheettmp.Columns["AN:AN"].ColumnWidth = colWith[39]; } catch (Exception) { }
                            try { sheettmp.Columns["AO:AO"].ColumnWidth = colWith[40]; } catch (Exception) { }
                            try { sheettmp.Columns["AP:AP"].ColumnWidth = colWith[41]; } catch (Exception) { }
                            try { sheettmp.Columns["AQ:AQ"].ColumnWidth = colWith[42]; } catch (Exception) { }
                            try { sheettmp.Columns["AR:AR"].ColumnWidth = colWith[43]; } catch (Exception) { }
                            try { sheettmp.Columns["AS:AS"].ColumnWidth = colWith[44]; } catch (Exception) { }
                            try { sheettmp.Columns["AT:AT"].ColumnWidth = colWith[45]; } catch (Exception) { }
                            try { sheettmp.Columns["AU:AU"].ColumnWidth = colWith[46]; } catch (Exception) { }
                            try { sheettmp.Columns["AV:AV"].ColumnWidth = colWith[47]; } catch (Exception) { }
                            try { sheettmp.Columns["AW:AW"].ColumnWidth = colWith[48]; } catch (Exception) { }
                            try { sheettmp.Columns["AX:AX"].ColumnWidth = colWith[49]; } catch (Exception) { }
                            try { sheettmp.Columns["AY:AY"].ColumnWidth = colWith[50]; } catch (Exception) { }
                            try { sheettmp.Columns["AZ:AZ"].ColumnWidth = colWith[51]; } catch (Exception) { }
                            try { sheettmp.Columns["BA:BA"].ColumnWidth = colWith[52]; } catch (Exception) { }
                            try { sheettmp.Columns["BB:BB"].ColumnWidth = colWith[53]; } catch (Exception) { }
                            try { sheettmp.Columns["BC:BC"].ColumnWidth = colWith[54]; } catch (Exception) { }
                            try { sheettmp.Columns["BD:BD"].ColumnWidth = colWith[55]; } catch (Exception) { }
                            try { sheettmp.Columns["BE:BE"].ColumnWidth = colWith[56]; } catch (Exception) { }
                            try { sheettmp.Columns["BF:BF"].ColumnWidth = colWith[57]; } catch (Exception) { }
                            try { sheettmp.Columns["BG:BG"].ColumnWidth = colWith[58]; } catch (Exception) { }
                            try { sheettmp.Columns["BH:BH"].ColumnWidth = colWith[59]; } catch (Exception) { }
                            try { sheettmp.Columns["BI:BI"].ColumnWidth = colWith[60]; } catch (Exception) { }
                            try { sheettmp.Columns["BG:BG"].ColumnWidth = colWith[61]; } catch (Exception) { }
                            try { sheettmp.Columns["BK:BK"].ColumnWidth = colWith[62]; } catch (Exception) { }
                            try { sheettmp.Columns["BL:BL"].ColumnWidth = colWith[63]; } catch (Exception) { }
                            try { sheettmp.Columns["BM:BM"].ColumnWidth = colWith[64]; } catch (Exception) { }
                            try { sheettmp.Columns["BN:BN"].ColumnWidth = colWith[65]; } catch (Exception) { }
                            try { sheettmp.Columns["BO:BO"].ColumnWidth = colWith[66]; } catch (Exception) { }
                            try { sheettmp.Columns["BP:BP"].ColumnWidth = colWith[67]; } catch (Exception) { }
                            try { sheettmp.Columns["BQ:BQ"].ColumnWidth = colWith[68]; } catch (Exception) { }
                            try { sheettmp.Columns["BR:BR"].ColumnWidth = colWith[69]; } catch (Exception) { }
                            try { sheettmp.Columns["BS:BS"].ColumnWidth = colWith[70]; } catch (Exception) { }
                            try { sheettmp.Columns["BT:BT"].ColumnWidth = colWith[71]; } catch (Exception) { }
                            try { sheettmp.Columns["BU:BU"].ColumnWidth = colWith[72]; } catch (Exception) { }
                            try { sheettmp.Columns["BV:BV"].ColumnWidth = colWith[73]; } catch (Exception) { }
                            try { sheettmp.Columns["BW:BW"].ColumnWidth = colWith[74]; } catch (Exception) { }
                            try { sheettmp.Columns["BX:BX"].ColumnWidth = colWith[75]; } catch (Exception) { }
                            try { sheettmp.Columns["BY:BY"].ColumnWidth = colWith[76]; } catch (Exception) { }
                            try { sheettmp.Columns["BZ:BZ"].ColumnWidth = colWith[77]; } catch (Exception) { }
                            try { sheettmp.Columns["CA:CA"].ColumnWidth = colWith[78]; } catch (Exception) { }
                            try { sheettmp.Columns["CB:CB"].ColumnWidth = colWith[79]; } catch (Exception) { }
                            try { sheettmp.Columns["CC:CC"].ColumnWidth = colWith[80]; } catch (Exception) { }
                            try { sheettmp.Columns["CD:CD"].ColumnWidth = colWith[81]; } catch (Exception) { }
                            try { sheettmp.Columns["CE:CE"].ColumnWidth = colWith[82]; } catch (Exception) { }
                            try { sheettmp.Columns["CF:CF"].ColumnWidth = colWith[83]; } catch (Exception) { }
                            try { sheettmp.Columns["CG:CG"].ColumnWidth = colWith[84]; } catch (Exception) { }
                            try { sheettmp.Columns["CH:CH"].ColumnWidth = colWith[85]; } catch (Exception) { }
                            try { sheettmp.Columns["CI:CI"].ColumnWidth = colWith[86]; } catch (Exception) { }
                            try { sheettmp.Columns["CG:CG"].ColumnWidth = colWith[87]; } catch (Exception) { }
                            try { sheettmp.Columns["CK:CK"].ColumnWidth = colWith[88]; } catch (Exception) { }
                            try { sheettmp.Columns["CL:CL"].ColumnWidth = colWith[89]; } catch (Exception) { }
                            try { sheettmp.Columns["CM:CM"].ColumnWidth = colWith[90]; } catch (Exception) { }
                            try { sheettmp.Columns["CN:CN"].ColumnWidth = colWith[91]; } catch (Exception) { }
                            try { sheettmp.Columns["CO:CO"].ColumnWidth = colWith[92]; } catch (Exception) { }
                            try { sheettmp.Columns["CP:CP"].ColumnWidth = colWith[93]; } catch (Exception) { }
                            try { sheettmp.Columns["CQ:CQ"].ColumnWidth = colWith[94]; } catch (Exception) { }
                            try { sheettmp.Columns["CR:CR"].ColumnWidth = colWith[95]; } catch (Exception) { }
                            try { sheettmp.Columns["CS:CS"].ColumnWidth = colWith[96]; } catch (Exception) { }
                            try { sheettmp.Columns["CT:CT"].ColumnWidth = colWith[97]; } catch (Exception) { }
                            try { sheettmp.Columns["CU:CU"].ColumnWidth = colWith[98]; } catch (Exception) { }
                            try { sheettmp.Columns["CV:CV"].ColumnWidth = colWith[99]; } catch (Exception) { }
                            try { sheettmp.Columns["CW:CW"].ColumnWidth = colWith[100]; } catch (Exception) { }
                            try { sheettmp.Columns["CX:CX"].ColumnWidth = colWith[101]; } catch (Exception) { }
                            try { sheettmp.Columns["CY:CY"].ColumnWidth = colWith[102]; } catch (Exception) { }
                            try { sheettmp.Columns["CZ:CZ"].ColumnWidth = colWith[103]; } catch (Exception) { }
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
