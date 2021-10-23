using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Reflection;
using System.Collections;

namespace WordDotx
{
    /// <summary>
    /// Класс для создание сервераобработки документов
    /// </summary>
    public class FarmWordDotx : Lib.TaskWordBase.FarmWordDotxBase
    {
        #region  Private Param
        /// <summary>
        /// Внутренний объект нашего Сервера когда сервер должен быть единственным в нашем приложении
        /// </summary>
        private static WordDotxServer _CurrentWordDotxServer;

        /// <summary>
        /// Папка по умолчанию для нашего файла с источником шаблонов
        /// </summary>
        private static string _DefaultPathSource;

        /// <summary>
        /// Папка по умолчанию для нашего файла в который положим результат
        /// </summary>
        private static string _DefaultPathTarget;

        /// <summary>
        /// Поведение по умолчанию нужно заменить файл или нет
        /// </summary>
        private static bool _DefReplaseFileTarget;
        #endregion

        #region  Public Param
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
        /// Получаем версию нашего приложения в дальнейшем если будут меняться классы  и отражают следующие изменения
        ///    значение с индексом  0 меняют интерфейсы пользователя старые версии могут не работать совсем
        ///    значение с индексом  1 добовляется функционал но это новый функционал и на пользователя вообще не влияет 
        ///    значение с индексом  2 изменяется текущий функционал но для пользователя изменений нет
        ///    значение с индексом  3 если изменения структуры вообще не меняются а только правится ошибка
        /// </summary>
        public static int[] VersionDll
        {
            get
            {
                int[] rez = { 1, 0, 0, 1 };
                string ss = Assembly.GetExecutingAssembly().GetName().Version.ToString();
                string[] tmp = ss.Split('.');

                for (int i = 0; i < rez.Length; i++)
                {
                    rez[i] = int.Parse(tmp[i]);
                }

                return rez;
            }
            private set { }
        }

        /// <summary>
        /// Папка по умолчанию для нашего файла с источником шаблонов
        /// </summary>
        public static string DefaultPathSource
        {
            get
            {
                return _DefaultPathSource;
            }
            private set { }
        }

        /// <summary>
        /// Папка по умолчанию для нашего файла в который положим результат
        /// </summary>
        public static string DefaultPathTarget
        {
            get
            {
                return _DefaultPathTarget;
            }
            private set { }
        }

        /// <summary>
        /// Поведение по умолчанию нужно заменить файл или нет
        /// </summary>
        public static bool DefReplaseFileTarget
        {
            get
            {
                return _DefReplaseFileTarget;
            }
            private set { }
        }

        /// <summary>
        /// Пул который обеспечивает парралельную обработку сразу нескольких заданий
        /// </summary>
        public static WorkerList PoolWorkerList = new WorkerList();

        #endregion

        #region  Public Method
        /// <summary>
        /// Создание сервера который будет обрабатывать наши объекты ворда в эдиничном экземпляре
        /// </summary>
        /// <param name="DefaultPathSource">Папка по умолчанию для нашего файла с источником шаблонов</param>
        /// <param name="DefaultPathTarget">Папка по умолчанию для нашего файла в который положим результат</param>
        /// <param name="DefReplaseFileTarget">Поведение по умолчанию нужно заменить файл или нет</param>
        /// <returns>Возвращет наш сервер который будет обрабатывать отчёты</returns>
        public static WordDotxServer CreateWordDotxServer(string DefaultPathSource, string DefaultPathTarget, bool DefReplaseFileTarget)
        {
            try
            {
                if (_CurrentWordDotxServer == null)
                {
                    _CurrentWordDotxServer = new WordDotxServer(DefaultPathSource, DefaultPathTarget, DefReplaseFileTarget);
                    _DefaultPathSource = CurrentWordDotxServer.DefaultPathSource;
                    _DefaultPathTarget = CurrentWordDotxServer.DefaultPathTarget;
                    _DefReplaseFileTarget = CurrentWordDotxServer.DefReplaseFileTarget;
                }
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
        /// <param name="DefaultPathSource">Папка по умолчанию для нашего файла с источником шаблонов</param>
        /// <param name="DefaultPathTarget">Папка по умолчанию для нашего файла в который положим результат</param>
        /// <returns>Возвращет наш сервер который будет обрабатывать отчёты</returns>
        public static WordDotxServer CreateWordDotxServer(string DefaultPathSource, string DefaultPathTarget)
        {
            try
            {
                if (_CurrentWordDotxServer == null)
                {
                    _CurrentWordDotxServer = new WordDotxServer(DefaultPathSource, DefaultPathTarget, true);
                    _DefaultPathSource = CurrentWordDotxServer.DefaultPathSource;
                    _DefaultPathTarget = CurrentWordDotxServer.DefaultPathTarget;
                    _DefReplaseFileTarget = CurrentWordDotxServer.DefReplaseFileTarget;
                }
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
                if (_CurrentWordDotxServer == null)
                {
                    _CurrentWordDotxServer = new WordDotxServer(DefPathSorsAndTarget);
                    _DefaultPathSource = CurrentWordDotxServer.DefaultPathSource;
                    _DefaultPathTarget = CurrentWordDotxServer.DefaultPathTarget;
                    _DefReplaseFileTarget = CurrentWordDotxServer.DefReplaseFileTarget;
                }
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
                if (_CurrentWordDotxServer == null)
                {
                    _CurrentWordDotxServer = new WordDotxServer();
                    _DefaultPathSource = CurrentWordDotxServer.DefaultPathSource;
                    _DefaultPathTarget = CurrentWordDotxServer.DefaultPathTarget;
                    _DefReplaseFileTarget = CurrentWordDotxServer.DefReplaseFileTarget;
                }
                return _CurrentWordDotxServer;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(string.Format("{0}.FarmWordDotx   Упали с ошибкой при создании сервера: ({1})", "FarmWordDotx", ex.Message));
            }
        }

        
        #endregion

    }
}
