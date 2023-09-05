using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Waher.Content;
using Waher.Runtime.Inventory;

namespace TAG.Content.Microsoft
{
	/// <summary>
	/// Decodes Excel documents
	/// </summary>
	public class ExcelDecoder : IContentDecoder
	{
		/// <summary>
		/// Decodes Excel documents
		/// </summary>
		public ExcelDecoder()
		{
		}

		/// <summary>
		/// application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
		/// </summary>
		public const string ExcelDocumentContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

		/// <summary>
		/// docx
		/// </summary>
		public const string ExcelDocumentExtension = "xlsx";

		/// <summary>
		/// Content-Types supported by the decoder.
		/// </summary>
		public string[] ContentTypes => new string[] { ExcelDocumentContentType };

		/// <summary>
		/// File extensions supported by the decoder.
		/// </summary>
		public string[] FileExtensions => new string[] { ExcelDocumentExtension };

		/// <summary>
		/// How well the decoder handles content of a given type.
		/// </summary>
		/// <param name="ContentType">Content-Type</param>
		/// <param name="Grade">How well the Content-Type is supported.</param>
		/// <returns>If the decoder supports the given content type.</returns>
		public bool Decodes(string ContentType, out Grade Grade)
		{
			if (string.Compare(ContentType, ExcelDocumentContentType, true) == 0)
			{
				Grade = Grade.Ok;
				return true;
			}
			else
			{
				Grade = Grade.NotAtAll;
				return false;
			}
		}

		/// <summary>
		/// Decodes an encoded object.
		/// </summary>
		/// <param name="ContentType">Content-Type of encoded object.</param>
		/// <param name="Data">Binary data</param>
		/// <param name="Encoding">Any default encoding provided.</param>
		/// <param name="Fields">Fields available in request.</param>
		/// <param name="BaseUri">Base URI of object.</param>
		/// <returns>Decoded object.</returns>
		public Task<object> DecodeAsync(string ContentType, byte[] Data, Encoding Encoding, KeyValuePair<string, string>[] Fields, Uri BaseUri)
		{
			MemoryStream ms = new MemoryStream(Data);
			SpreadsheetDocument Doc = SpreadsheetDocument.Open(ms, false);

			// Note: Do not dispose MemoryStream. The document needs the stream to remain open.
			//       This incurrs no memory loss while using only the MemoryStream, as no
			//       unmanaged resources are used. The GC will reclaim unused memory once
			//       no longer using the document.
				
			return Task.FromResult<object>(Doc);
		}

		/// <summary>
		/// Tries to get the Content-Type given a file extension.
		/// </summary>
		/// <param name="FileExtension">File extension.</param>
		/// <param name="ContentType">Content-Type, if recognized.</param>
		/// <returns>If file extension was recognized.</returns>
		public bool TryGetContentType(string FileExtension, out string ContentType)
		{
			if (string.Compare(FileExtension, ExcelDocumentExtension, true) == 0)
			{
				ContentType = ExcelDocumentContentType;
				return true;
			}
			else
			{
				ContentType = null;
				return false;
			}
		}

		/// <summary>
		/// Tries to get the Content-Type given a file extension.
		/// </summary>
		/// <param name="FileExtension">File extension.</param>
		/// <param name="ContentType">Content-Type, if recognized.</param>
		/// <returns>If file extension was recognized.</returns>
		public bool TryGetFileExtension(string ContentType, out string FileExtension)
		{
			if (string.Compare(ContentType, ExcelDocumentContentType, true) == 0)
			{
				FileExtension = ExcelDocumentExtension;
				return true;
			}
			else
			{
				FileExtension = null;
				return false;
			}
		}
	}
}
