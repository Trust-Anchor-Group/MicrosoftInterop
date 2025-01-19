using DocumentFormat.OpenXml.Packaging;
using System.Text;
using System.Threading.Tasks;
using TAG.Content.Microsoft;
using TAG.Content.Microsoft.Content;
using Waher.Content;
using Waher.Content.Markdown;
using Waher.Networking.HTTP;

namespace TAG.Service.MicrosoftInterop.WebServices
{
    /// <summary>
    /// Converts a Word document to Markdown
    /// </summary>
    public class WordToMarkdown : HttpSynchronousResource, IHttpPostMethod
	{
		private readonly HttpAuthenticationScheme[] authenticationSchemes;

		/// <summary>
		/// Converts a Word document to Markdown
		/// </summary>
		/// <param name="AuthenticationSchemes">Authentication schemes.</param>
		public WordToMarkdown(params HttpAuthenticationScheme[] AuthenticationSchemes)
			: base("/MicrosoftInterop/WordToMarkdown")
		{
			this.authenticationSchemes = AuthenticationSchemes;
		}

		/// <summary>
		/// If sub-paths are handled.
		/// </summary>
		public override bool HandlesSubPaths => false;

		/// <summary>
		/// If User sessions are required
		/// </summary>
		public override bool UserSessions => true;

		/// <summary>
		/// Gets available authentication schemes
		/// </summary>
		/// <param name="Request">Request object.</param>
		/// <returns>Array of authentication schemes.</returns>
		public override HttpAuthenticationScheme[] GetAuthenticationSchemes(HttpRequest Request)
		{
			return this.authenticationSchemes;
		}

		/// <summary>
		/// If the POST method is supported.
		/// </summary>
		public bool AllowsPOST => true;

		/// <summary>
		/// Executes the POST method
		/// </summary>
		/// <param name="Request">Request object.</param>
		/// <param name="Response">Response object.</param>
		public async Task POST(HttpRequest Request, HttpResponse Response)
		{
			if (!Request.HasData)
			{
				await Response.SendResponse(new BadRequestException("No content."));
				return;
			}

			ContentResponse Decoded = await Request.DecodeDataAsync();
			if (Decoded.HasError)
			{
				await Response.SendResponse(Decoded.Error);
				return;
			}

			if (!(Decoded.Decoded is WordprocessingDocument Doc))
			{
				await Response.SendResponse(new BadRequestException("Content not a Word document (.docx)."));
				return;
			}

			string Markdown = WordUtilities.ExtractAsMarkdown(Doc, string.Empty, out _);
			byte[] Data = WordToMarkdownConverter.Utf8WithBOM.GetBytes(Markdown);

			Response.ContentType = MarkdownCodec.ContentType + "; charset=utf-8";
			await Response.Write(true, Data);
		}
	}
}
