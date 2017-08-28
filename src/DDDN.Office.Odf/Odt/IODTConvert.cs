using System;

namespace DDDN.Office.Odf.Odt
{
    public interface IODTConvert : IDisposable
    {
        ODTConvertData Convert();
    }
}