using System;
using System.IO;

namespace HtmlToWord.Service
{
    public interface IWordApplication : IDisposable
    {
        void SetDocumentSize(float width, float height);
        bool ConvertToWord(FileInfo htmlFile, FileInfo docFileInfo, out string message);
    }
}