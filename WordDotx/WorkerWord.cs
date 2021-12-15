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
    public class WorkerWord : Lib.WorkerWordBase
    {
        /// <summary>
        /// Ссылка на задание по которому идёт расчёт результата
        /// </summary>
        public new TaskWord TaskWrk
        {
            get
            {
                return base.TaskWrk;
            }
            private set { }
        }

        /// <summary>
        /// Событие исключения которое возникло в работнике и он не может продолжать обрабатывать документы
        /// </summary>
        public event EventHandler<EvWorkerWordError> onEvWorkerError;

        /// <summary>
        /// Конструктор
        /// </summary>
        public WorkerWord()
        {
            base.onEvWorkerBaseError += Worker_onEvWorkerBaseError;
        }

        /// <summary>
        /// Обработка исключения в базовом классе
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Worker_onEvWorkerBaseError(object sender, Lib.EvWorkerWordBaseError e)
        {
            try
            {
                if (this.onEvWorkerError != null)
                {
                    EvWorkerWordError ArgErrorW = new EvWorkerWordError(this, e.ErrorMessage);
                    this.onEvWorkerError.Invoke(this, ArgErrorW);
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.Worker_onEvenWorkerBaseError   Упали с ошибкой в потоке которй обслуживает сервис по формированию отчёта: ({1})", this.GetType().Name, ex.Message));
            }
        }
    }
}
