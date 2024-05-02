using System.IO;
using System.Threading.Tasks;
using TAG.Content.Microsoft.Content;
using Waher.Content;
using Waher.Content.Html.JavaScript;
using Waher.IoTGateway;
using Waher.Networking.HTTP;

namespace TAG.Service.MicrosoftInterop.WebServices
{
    /// <summary>
    /// Appends information to the Prompt Javascript file, to allow for
    /// Excel document uploads.
    /// </summary>
    public class AppendingPromptJs : HttpSynchronousResource, IHttpGetMethod
	{
		private readonly HttpAuthenticationScheme[] authenticationSchemes;

		/// <summary>
		/// Appends information to the Prompt Javascript file, to allow for
		/// Excel document uploads.
		/// </summary>
		public AppendingPromptJs(params HttpAuthenticationScheme[] AuthenticationSchemes)
			: base("/Prompt.js")
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
			string FileName1 = Path.Combine(Gateway.RootFolder, "Prompt.js");
			string Javascript1 = await Resources.ReadAllTextAsync(FileName1);

			string FileName2 = Path.Combine(Gateway.RootFolder, "MicrosoftInterop", "PromptAddendum.js");
			string Javascript2 = await Resources.ReadAllTextAsync(FileName2);

			Javascript1 += Javascript2;

			Response.ContentType = JavaScriptCodec.JavaScriptContentTypes[0] + "; charset=utf-8";
			await Response.Write(WordToMarkdownConverter.Utf8WithBOM.GetBytes(Javascript1));
		}
	}
}
