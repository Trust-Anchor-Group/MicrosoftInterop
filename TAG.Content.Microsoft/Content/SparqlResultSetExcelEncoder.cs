using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Waher.Content;
using Waher.Content.Semantic;
using Waher.Runtime.Inventory;
using Waher.Script.Abstraction.Elements;
using Waher.Script.Objects.Matrices;

namespace TAG.Content.Microsoft.Content
{
	/// <summary>
	/// Encoder of semantic information from SPARQL queries to an Excel spreadsheet.
	/// </summary>
	public class SparqlResultSetExcelEncoder : IContentEncoder
	{
		/// <summary>
		/// Encoder of semantic information from SPARQL queries to an Excel spreadsheet.
		/// </summary>
		public SparqlResultSetExcelEncoder()
		{
		}

		/// <summary>
		/// Supported Internet Content Types.
		/// </summary>
		public string[] ContentTypes => SparqlResultSetContentTypes;

		private static readonly string[] SparqlResultSetContentTypes = new string[]
		{
			ExcelDecoder.ExcelDocumentContentType
		};

		/// <summary>
		/// Supported file extensions.
		/// </summary>
		public string[] FileExtensions => SparqlResultSetFileExtensions;

		private static readonly string[] SparqlResultSetFileExtensions = new string[]
		{
			"xlsx"
		};

		/// <summary>
		/// If the encoder encodes a specific object.
		/// </summary>
		/// <param name="Object">Object to encode.</param>
		/// <param name="Grade">How well the encoder supports the given object.</param>
		/// <param name="AcceptedContentTypes">Accepted content types.</param>
		/// <returns>If the encoder encodes the given object.</returns>
		public bool Encodes(object Object, out Grade Grade, params string[] AcceptedContentTypes)
		{
			if (Object is SparqlResultSet &&
				InternetContent.IsAccepted(SparqlResultSetContentTypes, AcceptedContentTypes))
			{
				Grade = Grade.Excellent;
				return true;
			}
			else if (Object is ObjectMatrix M && M.HasColumnNames &&
				InternetContent.IsAccepted(SparqlResultSetContentTypes, AcceptedContentTypes))
			{
				Grade = Grade.Ok;
				return true;
			}
			else if (Object is bool &&
				InternetContent.IsAccepted(SparqlResultSetContentTypes, AcceptedContentTypes))
			{
				Grade = Grade.Barely;
				return true;
			}
			else
			{
				Grade = Grade.NotAtAll;
				return false;
			}
		}

		/// <summary>
		/// Encodes an object
		/// </summary>
		/// <param name="Object">Object to encode</param>
		/// <param name="Encoding">Encoding</param>
		/// <param name="Progress">Optional progress reporting of encoding/decoding. Can be null.</param>
		/// <param name="AcceptedContentTypes">Accepted content types.</param>
		/// <returns>Encoded object.</returns>
		public Task<ContentResponse> EncodeAsync(object Object, Encoding Encoding, ICodecProgress Progress, params string[] AcceptedContentTypes)
		{
			byte[] Bin;

			if (Object is SparqlResultSet Result)
				Bin = Encode(Result);
			else if (Object is IMatrix M)
				Bin = Encode(M);
			else if (Object is bool b)
				Bin = Encode(b);
			else
				return Task.FromResult(new ContentResponse(new ArgumentException("Unable to encode object.", nameof(Object))));

			return Task.FromResult(new ContentResponse(ExcelDecoder.ExcelDocumentContentType, Object, Bin));
		}

		private static byte[] Encode(SparqlResultSet Result)
		{
			string FilePath = Path.GetTempFileName();
			ExcelUtilities.ConvertResultSetToExcel(Result, FilePath, "Result");
			byte[] Bin = File.ReadAllBytes(FilePath);
			File.Delete(FilePath);

			return Bin;
		}

		private static byte[] Encode(IMatrix Result)
		{
			string FilePath = Path.GetTempFileName();
			ExcelUtilities.ConvertMatrixToExcel(Result, FilePath, "Result");
			byte[] Bin = File.ReadAllBytes(FilePath);
			File.Delete(FilePath);

			return Bin;
		}

		private static byte[] Encode(bool Result)
		{
			SparqlResultSet ResultSet = new SparqlResultSet(Result);
			return Encode(ResultSet);
		}

		/// <summary>
		/// Tries to get the content type of content of a given file extension.
		/// </summary>
		/// <param name="FileExtension">File Extension</param>
		/// <param name="ContentType">Content Type, if recognized.</param>
		/// <returns>If File Extension was recognized and Content Type found.</returns>
		public bool TryGetContentType(string FileExtension, out string ContentType)
		{
			if (string.Compare(FileExtension, SparqlResultSetFileExtensions[0], true) == 0)
			{
				ContentType = SparqlResultSetContentTypes[0];
				return true;
			}
			else
			{
				ContentType = null;
				return false;
			}
		}

		/// <summary>
		/// Tries to get the file extension of content of a given content type.
		/// </summary>
		/// <param name="ContentType">Content Type</param>
		/// <param name="FileExtension">File Extension, if recognized.</param>
		/// <returns>If Content Type was recognized and File Extension found.</returns>
		public bool TryGetFileExtension(string ContentType, out string FileExtension)
		{
			if (Array.IndexOf(SparqlResultSetContentTypes, ContentType) >= 0)
			{
				FileExtension = SparqlResultSetFileExtensions[0];
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
