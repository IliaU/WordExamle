using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDotx
{
    /// <summary>
    /// Базовый класс для заданий которые могут выполнять задания из очереди асинхронно
    /// </summary>
    public class WorkerExcel : Lib.WorkerExcelBase
    {
        /// <summary>
        /// Ссылка на задание по которому идёт расчёт результата
        /// </summary>
        public new TaskExcel TaskExl
        {
            get
            {
                return base.TaskExl;
            }
            private set { }
        }

        /// <summary>
        /// Событие исключения которое возникло в работнике и он не может продолжать обрабатывать документы
        /// </summary>
        public event EventHandler<EvWorkerExcelError> onEvWorkerExcelError;

        /// <summary>
        /// Конструктор
        /// </summary>
        public WorkerExcel()
        {
            base.onEvWorkerExcelBaseError += Worker_onEvWorkerBaseError;
 
        }

        /// <summary>
        /// Обработка исключения в базовом классе
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Worker_onEvWorkerBaseError(object sender, Lib.EvWorkerExcelBaseError e)
        {
            try
            {
                if (this.onEvWorkerExcelError != null)
                {
                    EvWorkerExcelError ArgErrorW = new EvWorkerExcelError(this, e.ErrorMessage);
                    this.onEvWorkerExcelError.Invoke(this, ArgErrorW);
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.Worker_onEvenWorkerBaseError   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
            }
        }
    }
}
