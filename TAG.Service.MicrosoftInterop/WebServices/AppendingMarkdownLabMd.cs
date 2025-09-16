using System.IO;
using System.Threading.Tasks;
using TAG.Content.Microsoft.Content;
using Waher.Content;
using Waher.Content.Markdown.Web;
using Waher.IoTGateway;
using Waher.Networking.HTTP;
using Waher.Runtime.IO;

namespace TAG.Service.MicrosoftInterop.WebServices
{
    /// <summary>
    /// Appends information to the Markdown Lab Markdown page, to allow for
    /// Word document uploads.
    /// </summary>
    public class AppendingMarkdownLabMd : HttpSynchronousResource, IHttpGetMethod
	{
		private readonly HttpAuthenticationScheme[] authenticationSchemes;

		/// <summary>
		/// Appends information to the Markdown Lab Markdown page, to allow for
		/// Word document uploads.
		/// </summary>
		public AppendingMarkdownLabMd(params HttpAuthenticationScheme[] AuthenticationSchemes)
			: base("/MarkdownLab/MarkdownLab.md")
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
		/// If the GET method is supported.
		/// </summary>
		public bool AllowsGET => true;

		/// <summary>
		/// Executes the POST method
		/// </summary>
		/// <param name="Request">Request object.</param>
		/// <param name="Response">Response object.</param>
		public async Task GET(HttpRequest Request, HttpResponse Response)
		{
			string FileName1 = Path.Combine(Gateway.RootFolder, "MarkdownLab", "MarkdownLab.md");
			string Markdown1 = await Files.ReadAllTextAsync(FileName1);
			int i = Markdown1.IndexOf("</section>");

			if (i >= 0)
			{
				string FileName2 = Path.Combine(Gateway.RootFolder, "MicrosoftInterop", "MarkdownLabAddendum.md");
				string Markdown2 = await Files.ReadAllTextAsync(FileName2);

				Markdown1 = Markdown1.Insert(i, Markdown2);
			}

			await SendMarkdownAsHtml(Request, Response, Markdown1, FileName1);
		}

		/// <summary>
		/// Converts Markdown to HTML and returns it to the client.
		/// </summary>
		/// <param name="Request">Request object.</param>
		/// <param name="Response">Response object.</param>
		/// <param name="Markdown">Markdown to process.</param>
		/// <param name="FileName">Name of markdown file.</param>
		public static async Task SendMarkdownAsHtml(HttpRequest Request, HttpResponse Response,
			string Markdown, string FileName)
		{
			MarkdownToHtmlConverter Converter = new MarkdownToHtmlConverter();
			byte[] Bin = WordToMarkdownConverter.Utf8WithBOM.GetBytes(Markdown);
			
			using MemoryStream Input = new MemoryStream(Bin);
			using MemoryStream Output = new MemoryStream();
			
			string MarkdownContentType = Converter.FromContentTypes[0];
			string HtmlContentType = Converter.ToContentTypes[0];

			ConversionState State = new ConversionState(MarkdownContentType, Input, FileName,
				Request.Header.Resource, Request.Header.GetURL(), HtmlContentType, Output,
				Request.Session, Response.Progress, Request.Server, Request.TryGetLocalResourceFileName,
				Converter.ToContentTypes);

			await Converter.ConvertAsync(State);

			Bin = Output.ToArray();

			Response.ContentType = State.ToContentType;
			await Response.Write(true, Bin);
		}
	}
}
