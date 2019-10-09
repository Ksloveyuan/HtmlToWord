using System;
using System.IO;
using HtmlToWord.Core;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace HtmlToWord.Service
{
    public class WordApplication : IWordApplication
    {
        private readonly ILogger _logger;
        private readonly Application _word;

        private float _documentHeight;
        private float _documentWidth;

        public WordApplication(ILogger logger)
        {
            this._word = new Application {Visible = false};
            this._logger = logger;
        }

        public void Dispose()
        {
            try
            {
                this._word?.Quit();
            }
            catch (Exception e)
            {
                this._logger.Error("Failed to dispose, {0}", e);
            }
        }

        public void SetDocumentSize(float width, float height)
        {
            this._documentWidth = width;
            this._documentHeight = height;
        }

        public bool ConvertToWord(FileInfo htmlFile, FileInfo docFileInfo, out string message)
        {
            try
            {
                var doc = this._word.Documents.Open(htmlFile.FullName, Format: WdOpenFormat.wdOpenFormatWebPages,
                    ReadOnly: false);
                if (doc == null)
                {
                    message = $"Failed to export word of {htmlFile.Name}";
                    this._logger.Info(message);
                    return false;
                }

                this._logger.Debug("InlineShapes:{0}", doc.InlineShapes.Count);
                this._logger.Debug("Document width:{0}, height:{1}", this._documentWidth, this._documentHeight);
                foreach (InlineShape s in doc.InlineShapes)
                {
                    var inlineShape = s;
                    this._logger.Debug("type:{0} width:{1:f} height:{2:f}", inlineShape.Type, inlineShape.Width,
                        inlineShape.Height);

                    if (inlineShape.Type != WdInlineShapeType.wdInlineShapePicture &&
                        inlineShape.Type != WdInlineShapeType.wdInlineShapeLinkedPicture)
                    {
                        continue;
                    }

                    if (inlineShape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture)
                    {
                        inlineShape.LinkFormat.SavePictureWithDocument = true;
                        inlineShape.LinkFormat.BreakLink();
                    }

                    if (inlineShape.Width > this._documentWidth)
                    {
                        inlineShape.LockAspectRatio = MsoTriState.msoTrue;
                        inlineShape.Width = this._documentWidth;
                        this._logger.Debug("resize by width, updated width:{0:f} height:{1:f}", inlineShape.Width,
                            inlineShape.Height);
                    }

                    if (inlineShape.Height > this._documentHeight)
                    {
                        inlineShape.LockAspectRatio = MsoTriState.msoTrue;
                        inlineShape.Height = this._documentHeight;
                        this._logger.Debug("resize by height, updated width:{0:f} height:{1:f}", inlineShape.Width,
                            inlineShape.Height);
                    }
                }

                doc.SaveAs2000(docFileInfo.FullName, WdSaveFormat.wdFormatDocumentDefault,
                    ReadOnlyRecommended: false);
                doc.Close();

                message = $"Export word of {docFileInfo.Name} successfully";
                this._logger.Info(message);
                return true;
            }
            catch (Exception e)
            {
                message = $"Failed to export word of {docFileInfo.Name}. Error is {e}";
                this._logger.Error(message);
                return false;
            }
        }
    }
}