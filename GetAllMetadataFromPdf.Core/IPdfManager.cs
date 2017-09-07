using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetAllMetadataFromPdf.Core
{
    public interface IPdfManager<T>
    {
        T CreateExcelWithPdfMetadata(T pdfPath);
        IDictionary<T,T> GetAllPdfMetadata(T pdfPath);
    }
}
