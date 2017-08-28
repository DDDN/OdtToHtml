using System;

namespace DDDN.Office.Odf.Odt
{
    public interface IODTConvert : IDisposable
    {
        string GetHtml();
        string GetCss();
    }
}