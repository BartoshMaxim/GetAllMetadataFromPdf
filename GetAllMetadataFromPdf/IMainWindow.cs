using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetAllMetadataFromPdf
{
    public interface IMainWindow<T>
    {
        T PdfPath { get; }
        event EventHandler ImportMetaDataToExcelClick;
    }
}
