using System.IO;
using System.Threading.Tasks;
using TAG.Content.Microsoft;
using Waher.Content;
using Waher.Content.Html.Css;
using Waher.Content.Html.Javascript;
using Waher.IoTGateway;
using Waher.Networking.HTTP;

namespace TAG.Service.MicrosoftInterop.WebServices
{
	/// <summary>
	/// Appends information to the Markdown Lab CSS file, to allow for
	/// Word document uploads.
	/// </summary>
	public class AppendingMarkdownLabCss : HttpSynchronousResource, IHttpGetMethod
	{
		private readonly HttpAuthenticationScheme[] authenticationSchemes;

		/// <summary>
		/// Appends information to the Markdown Lab CSS file, to allow for
		/// Word document uploads.
		/// </summary>
		public AppendingMarkdownLabCss(params HttpAuthenticationScheme[] AuthenticationSchemes)
			: base("/MarkdownLab/MarkdownLab.css")
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
			string FileName = Path.Combine(Gateway.RootFolder, "MarkdownLab", "MarkdownLab.Css");
			string Css = await Resources.ReadAllTextAsync(FileName);

			Css = Css.Replace("min-height:60vh;", "min-height:50vh;");

			Response.ContentType = CssCodec.CssContentTypes[0] + "; charset=utf-8";
			await Response.Write(WordToMarkdownConverter.Utf8WithBOM.GetBytes(Css));
		}
	}
}
