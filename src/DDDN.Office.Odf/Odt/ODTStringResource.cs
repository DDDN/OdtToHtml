/*
* DDDN.Office.Odf.Odt.ODTStringResource
* 
* Copyright(C) 2017 Lukasz Jaskiewicz
* Author: Lukasz Jaskiewicz (lukasz@jaskiewicz.de, devdone@outlook.com)
*
* This program is free software; you can redistribute it and/or modify it under the terms of the
* GNU General Public License as published by the Free Software Foundation; version 2 of the License.
*
* This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
* warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU General Public License for more details.
*
* You should have received a copy of the GNU General Public License along with this program; if not, write
* to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using DDDN.Logging.Messages;
using DDDN.Office.Odf.Odt;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace DDDN.CrossBlog.Blog.Localization.ODF
{
	public class ODTStringResource : IODTStringResource
	{
		private string ResourceKey;
		private string ResourceFolderPath;

		public ODTStringResource(string resourceKey, string resourceFolderPath)
		{
			if (string.IsNullOrWhiteSpace(resourceKey))
			{
				throw new ArgumentException(LogMsg.StrArgNullOrWhite, nameof(resourceKey));
			}

			if (string.IsNullOrWhiteSpace(resourceFolderPath))
			{
				throw new ArgumentException(LogMsg.StrArgNullOrWhite, nameof(resourceFolderPath));
			}

			ResourceKey = resourceKey;
			ResourceFolderPath = resourceFolderPath;
		}

		public Dictionary<string, string> GetStrings()
		{
			var ret = new Dictionary<string, string>();
			var lastPointIndex = ResourceKey.LastIndexOf('.');
			var odtFileName = ResourceKey.Substring(lastPointIndex + 1);
			var l10nPath = Path.Combine(
				ResourceFolderPath,
				ResourceKey.Remove(lastPointIndex).Replace('.', '\\'),
				"l10n");

			var resourcefileFullPaths = Directory.GetFiles(l10nPath, $"{odtFileName}.*");

			foreach (var fileFullPath in resourcefileFullPaths)
			{
				var cultureNameFromFileName = Path.GetFileNameWithoutExtension(fileFullPath).Split('.').Last();

				using (IODTFile odtFile = new ODTFile(fileFullPath))
				{
					XDocument contentXDoc = odtFile.GetZipArchiveEntryAsXDocument("content.xml");

					var contentEle = contentXDoc.Root
						.Elements(XName.Get("body", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"))
						.Elements(XName.Get("text", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"))
						.First();

					foreach (var table in contentEle.Elements()
						.Where(p => p.Name.LocalName.Equals("table", StringComparison.CurrentCultureIgnoreCase)))
					{
						foreach (var row in table.Elements()
						.Where(p => p.Name.LocalName.Equals("table-row", StringComparison.CurrentCultureIgnoreCase))
						.Skip(1))
						{
							var cells = row.Elements()
								.Where(p => p.Name.LocalName.Equals("table-cell", StringComparison.CurrentCultureIgnoreCase));
							var translationKey = cells.First().Value;
							var translation = cells.Skip(1).First().Value;
							ret.Add($"{translationKey}.{cultureNameFromFileName}", translation);
						}
					}
				}
			}

			return ret;
		}
	}
}
