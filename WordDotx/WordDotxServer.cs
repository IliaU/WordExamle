using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        }
        //
        /// <summary>
        /// Конструктор для создания сервера
        /// </summary>
        public WordDotxServer() : this(Environment.CurrentDirectory, Environment.CurrentDirectory)
        {
        }





    }
}
