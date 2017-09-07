using GetAllMetadataFromPdf.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace GetAllMetadataFromPdf
{
    public class MainPresenter
    {
        private readonly IMainWindow<string> _mainWindow;
        private readonly IPdfManager<string> _pdfManager;
        private readonly IMessageService _messageService;

        public MainPresenter(IMainWindow<string> mainWindow, IMessageService messageService)
        {
            _pdfManager = new PdfManager();

            _messageService = messageService;

            _mainWindow = mainWindow;
            _mainWindow.ImportMetaDataToExcelClick += _mainWindow_ImportMetaDataToExcelClick;
        }

        private void _mainWindow_ImportMetaDataToExcelClick(object sender, EventArgs e)
        {
            string pdfPath = _mainWindow.PdfPath;
            if (File.Exists(pdfPath))
            {
                string excelFilePath = string.Empty;
                excelFilePath = _pdfManager.CreateExcelWithPdfMetadata(_mainWindow.PdfPath);
                _messageService.ShowMessage("MetaData from pdf was wrote to excel file: " + excelFilePath);
            }
            else if (pdfPath == string.Empty)
                _messageService.ShowError("Enter pdf path");
            else
                _messageService.ShowError("File not fount" + _mainWindow.PdfPath);
        }
    }
}
