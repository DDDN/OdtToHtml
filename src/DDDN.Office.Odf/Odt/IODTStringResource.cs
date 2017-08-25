using System.Collections.Generic;

namespace DDDN.CrossBlog.Blog.Localization.ODF
{
	public interface IODTStringResource
	{
		Dictionary<string, string> GetStrings();
	}
}