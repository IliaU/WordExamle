using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Word;

namespace WordDotx
{
    /// <summary>
    /// Класс для создание сервераобработки документов
    /// </summary>
    public class FarmWordDotx
    {
        /// <summary>
        /// Внутренний объект нашего Сервера когда сервер должен быть единственным в нашем приложении
        /// </summary>
        private static WordDotxServer _CurrentWordDotxServer;

        /// <summary>
        /// Текущий  объект нашего Сервера когда сервер должен быть единственным в нашем приложении
        /// </summary>
        public static WordDotxServer CurrentWordDotxServer
        {
            get
            {
                return _CurrentWordDotxServer;
            }
            private set { }
        }

        /// <summary>
        /// Создание сервера который будет обрабатывать наши объекты ворда в эдиничном экземпляре
        /// </summary>
        /// <param name="DefaultPathSource">Папка по умолчанию для нашего файла с источником шаблонов</param>
        /// <param name="DefaultPathTarget">Папка по умолчанию для нашего файла в который положим результат</param>
        /// <returns>Возвращет наш сервер который будет обрабатывать отчёты</returns>
        public static WordDotxServer CreateWordDotxServer(string DefaultPathSource, string DefaultPathTarget)
        {
            try
            {
                if (_CurrentWordDotxServer == null) _CurrentWordDotxServer = new WordDotxServer(DefaultPathSource, DefaultPathTarget);
                return _CurrentWordDotxServer;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.FarmWordDotx   Упали с ошибкой при создании сервера: ({1})", "FarmWordDotx", ex.Message));
            }
        }

        /// <summary>
        /// Создание сервера который будет обрабатывать наши объекты ворда в эдиничном экземпляре
        /// </summary>
        /// <param name="DefPathSorsAndTarget">Если путь один и для входящих файлов и исходящих</param>
        /// <returns>Возвращет наш сервер который будет обрабатывать отчёты</returns>
        public static WordDotxServer CreateWordDotxServer(string DefPathSorsAndTarget)
        {
            try
            {
                if (_CurrentWordDotxServer == null) _CurrentWordDotxServer = new WordDotxServer(DefPathSorsAndTarget);
                return _CurrentWordDotxServer;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.FarmWordDotx   Упали с ошибкой при создании сервера: ({1})", "FarmWordDotx", ex.Message));
            }
        }

        /// <summary>
        /// Создание сервера который будет обрабатывать наши объекты ворда в эдиничном экземпляре
        /// </summary>
        /// <returns>Возвращет наш сервер который будет обрабатывать отчёты</returns>
        public static WordDotxServer CreateWordDotxServer()
        {
            try
            {
                if (_CurrentWordDotxServer == null) _CurrentWordDotxServer = new WordDotxServer();
                return _CurrentWordDotxServer;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.FarmWordDotx   Упали с ошибкой при создании сервера: ({1})", "FarmWordDotx", ex.Message));
            }
        }

    }
}
