using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Text;
using System.Threading.Tasks;
using Waher.Content;
using Waher.Content.Markdown;
using Waher.Runtime.Inventory;

namespace TAG.Content.Microsoft.Content
{
    /// <summary>
    /// Converts Word documents to Markdown documents
    /// </summary>
    public class WordToMarkdownConverter : IContentConverter
    {
        /// <summary>
        /// Converts Word documents
        /// </summary>
        public WordToMarkdownConverter()
        {
        }

        /// <summary>
        /// Content-Types from which the converter can convert.
        /// </summary>
        public string[] FromContentTypes => new string[] { WordDecoder.WordDocumentContentType };

        /// <summary>
        /// Content-Types to which the converter can convert.
        /// </summary>
        public string[] ToContentTypes => new string[] { MarkdownCodec.ContentType };

        /// <summary>
        /// How well conversion is established.
        /// </summary>
        public Grade ConversionGrade => Grade.Barely;

        /// <summary>
        /// Performs the actual conversion.
        /// </summary>
        /// <param name="State">State of the current conversion.</param>
        /// <returns>If the result is dynamic (true), or only depends on the source (false).</returns>
        public async Task<bool> ConvertAsync(ConversionState State)
        {
			using (WordprocessingDocument Doc = WordprocessingDocument.Open(State.From, false))
            {
                StringBuilder Markdown = new StringBuilder();
                WordUtilities.ExtractAsMarkdown(Doc, string.Empty, Markdown, out _);

                byte[] Data = Utf8WithBOM.GetBytes(Markdown.ToString());

                await State.To.WriteAsync(Data, 0, Data.Length);
                State.ToContentType += "; charset=utf-8";

                return false;
            }
        }

        /// <summary>
        /// UTF-8 encoding with BOM (byte-order mark).
        /// </summary>
        public static readonly Encoding Utf8WithBOM = new UTF8Encoding(true);
    }
}
