namespace HtmlToWord.Core
{
    public interface ILogger
    {
        void Info(string format, params object[] args);
        void Debug(string format, params object[] args);
        void Error(string format, params object[] args);
    }
}