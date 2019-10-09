using log4net;

namespace HtmlToWord.Core
{
    public class Logger : ILogger
    {
        private static readonly ILog InnerLogger = LogManager.GetLogger("example");

        public void Info(string format, params object[] args)
        {
            InnerLogger.InfoFormat(format, args);
        }
        public void Debug(string format, params object[] args)
        {
            InnerLogger.DebugFormat(format, args);
        }

        public void Error(string format, params object[] args)
        {
            InnerLogger.ErrorFormat(format, args);
        }
    }
}