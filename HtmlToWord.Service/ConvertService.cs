using System;
using System.Configuration;
using System.IO;
using System.ServiceModel;
using HtmlToWord.Contract;
using HtmlToWord.Core;

namespace HtmlToWord.Service
{
    [ServiceBehavior(ConcurrencyMode = ConcurrencyMode.Multiple, InstanceContextMode = InstanceContextMode.PerCall)]
    public class ConvertService : IConvert, IDisposable
    {
        private const string WordFolderName = "word";
        private const string HtmlFolderName = "html";

        private const string HtmlWrapper =
            "<!doctype html> <html lang=\"en\"> <head> <meta charset=\"UTF-8\"><title>Document</title> </head><body>{0}</body></html>";

        private static readonly string RootFolderPath;
        private static readonly int DocumentWidth;
        private static readonly int DocumentHeight;

        private readonly ILogger _logger;
        private readonly IWordApplication _word;

        static ConvertService()
        {
            RootFolderPath = ConfigurationManager.AppSettings["rootFolderPath"] ?? ".\\";

            if (int.TryParse(ConfigurationManager.AppSettings["documentWidth"], out var documentWidth))
            {
                DocumentWidth = documentWidth;
            }

            if (int.TryParse(ConfigurationManager.AppSettings["documentHeight"], out var documentHeight))
            {
                DocumentHeight = documentHeight;
            }

            if (!Directory.Exists(WordFolderPath))
            {
                Directory.CreateDirectory(WordFolderPath);
            }

            if (!Directory.Exists(HtmlFolderPath))
            {
                Directory.CreateDirectory(HtmlFolderPath);
            }
        }

        public ConvertService()
        {
            this._logger = new Logger();
            this._word = new WordApplication(this._logger);
            this._word.SetDocumentSize(DocumentWidth, DocumentHeight);
        }

        private static string WordFolderPath => Path.Combine(RootFolderPath, WordFolderName);
        private static string HtmlFolderPath => Path.Combine(RootFolderPath, HtmlFolderName);


        public CovertResult ToWord(string html)
        {
            var hash = html.GetHashCode().ToString("x8");
            this._logger.Info("Receive request: {0}", hash);

            var inputFileName = $"{hash}.html";
            var exportFileName = $"{hash}.doc";

            var inputFilePath = Path.Combine(HtmlFolderPath, inputFileName);
            var exportFilePath = Path.Combine(WordFolderPath, exportFileName);

            var inputFileInfo = new FileInfo(inputFilePath);
            var exportFileInfo = new FileInfo(exportFilePath);

            if (exportFileInfo.Exists)
            {
                this._logger.Info("Find cache for {0}, just return.", hash);
                return new CovertResult {Success = true, FileUrl = exportFileName};
            }

            try
            {
                var htmlFileContent = string.Format(HtmlWrapper, html);
                File.WriteAllText(inputFilePath, htmlFileContent);

                var success = this._word.ConvertToWord(inputFileInfo, exportFileInfo, out var message);
                return new CovertResult {Success = success, FileUrl = exportFileName, Message = message};
            }
            catch (Exception e)
            {
                this._logger.Info("Failed to export word of {0}", hash);
                this._logger.Error("Failed to export word of {0}. Error is {1}", hash, e);
                return new CovertResult {Success = false, Message = e.Message};
            }
        }

        public void Dispose()
        {
            this._word?.Dispose();
        }
    }
}