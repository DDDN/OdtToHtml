using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace DDDN.Office.Odf.Odt
{
	public interface IODTFile : IDisposable
	{
		XDocument GetZipArchiveEntryAsXDocument(string entryName);
		byte[] GetZipArchiveEntryFileContent(string entryPath);
	}
}