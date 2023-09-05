using System.IO;
using System.Threading.Tasks;
using TAG.Content.Microsoft;
using Waher.Content;
using Waher.Content.Html;
using Waher.Content.Markdown;
using Waher.IoTGateway;
using Waher.Networking.HTTP;

namespace TAG.Service.MicrosoftInterop.WebServices
{
	/// <summary>
	/// Appends information to the Prompt Markdown page, to allow for
	/// Excel document uploads.
	/// </summary>
	public class AppendingPromptMd : HttpSynchronousResource, IHttpGetMethod
	{
		private readonly HttpAuthenticationScheme[] authenticationSchemes;

		/// <summary>
		/// Appends information to the Prompt Markdown page, to allow for
		/// Excel document uploads.
		/// </summary>
		public AppendingPromptMd(params HttpAuthenticationScheme[] AuthenticationSchemes)
			: base("/Prompt.md")
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
			string FileName1 = Path.Combine(Gateway.RootFolder, "Prompt.md");
			string Markdown1 = await Resources.ReadAllTextAsync(FileName1);
			int i = Markdown1.IndexOf("\r\n\r\n=====================================================");

			if (i >= 0)
			{
				string FileName2 = Path.Combine(Gateway.RootFolder, "MicrosoftInterop", "PromptAddendum.md");
				string Markdown2 = await Resources.ReadAllTextAsync(FileName2);

				Markdown1 = Markdown1.Insert(i + 2, Markdown2);
			}

			MarkdownSettings Settings = new MarkdownSettings()
			{
				ResourceMap = Gateway.HttpServer,
				Variables = Request.Session
			};
			MarkdownDocument Doc = await MarkdownDocument.CreateAsync(Markdown1, Settings, string.Empty, 
				this.ResourceName, Gateway.GetUrl(this.ResourceName));

			string Html = await Doc.GenerateHTML();
			byte[] Bin = WordToMarkdownConverter.Utf8WithBOM.GetBytes(Html);

			Response.ContentType = HtmlCodec.HtmlContentTypes[0] + "; charset=utf-8";
			await Response.Write(Bin);
		}
	}
}
