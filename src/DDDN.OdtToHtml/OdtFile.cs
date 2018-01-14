/*
DDDN.OdtToHtml.OdtFile
Copyright(C) 2017-2018 Lukasz Jaskiewicz (lukasz@jaskiewicz.de)
- This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; version 2 of the License.
- This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
- You should have received a copy of the GNU General Public License along with this program; if not, write
to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;

namespace DDDN.OdtToHtml
{
	public class OdtFile : IOdtFile
	{
		private readonly ZipArchive ODTZipArchive;

		public OdtFile(string fileFullPath)
		{
			if (string.IsNullOrWhiteSpace(fileFullPath))
			{
				throw new ArgumentException(nameof(string.IsNullOrWhiteSpace), nameof(fileFullPath));
			}

			ODTZipArchive = ZipFile.OpenRead(fileFullPath);
		}

		public OdtFile(Stream fileStream)
		{
			if (fileStream == null)
			{
				throw new ArgumentNullException(nameof(fileStream));
			}

			ODTZipArchive = new ZipArchive(fileStream);
		}

		public IEnumerable<OdtEmbedContent> GetZipArchiveEntries()
		{
			var embedContentList = new List<OdtEmbedContent>();

			foreach (var zipEntry in ODTZipArchive.Entries)
			{
				using (var entryStream = zipEntry.Open())
				{
					using (var memStream = new MemoryStream())
					{
						entryStream.CopyTo(memStream);
						byte[] data = memStream.ToArray();

						var embedContent = new OdtEmbedContent
						{
							ContentFullName = zipEntry.FullName,
							Id = Guid.NewGuid(),
							Data = data
						};

						embedContentList.Add(embedContent);
					}
				}
			}

			return embedContentList;
		}

		public static XDocument GetZipArchiveEntryAsXDocument(byte[] zipEntryData)
		{
			if (zipEntryData == null)
			{
				return null;
			}

			XDocument contentXDoc = null;

			using (var stream = new MemoryStream(zipEntryData))
			{
				contentXDoc = XDocument.Load(stream);
			}

			return contentXDoc;
		}

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls

		protected virtual void Dispose(bool disposing)
		{
			if (!disposedValue)
			{
				if (disposing)
				{
					// TODO: dispose managed state (managed objects).
					ODTZipArchive.Dispose();
				}

				// TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
				// TODO: set large fields to null.

				disposedValue = true;
			}
		}

		// TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
		// ~ODTFile() {
		//   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
		//   Dispose(false);
		// }

		// This code added to correctly implement the disposable pattern.
		public void Dispose()
		{
			// Do not change this code. Put cleanup code in Dispose(bool disposing) above.
			Dispose(true);
			// TODO: uncomment the following line if the finalizer is overridden above.
			// GC.SuppressFinalize(this);
		}
		#endregion
	}
}
