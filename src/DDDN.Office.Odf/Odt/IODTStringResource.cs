using System.Collections.Generic;

namespace DDDN.Office.Odf.Odt
{
	public interface IODTStringResource
	{
		Dictionary<string, string> GetStrings();
	}
}