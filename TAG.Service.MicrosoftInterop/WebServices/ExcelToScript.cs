﻿using DocumentFormat.OpenXml.Packaging;
using System.Text;
using System.Threading.Tasks;
using TAG.Content.Microsoft;
using TAG.Content.Microsoft.Content;
using Waher.Content;
using Waher.Networking.HTTP;

namespace TAG.Service.MicrosoftInterop.WebServices
{
    /// <summary>
    /// Converts a Excel document to Script
    /// </summary>
    public class ExcelToScript : HttpSynchronousResource, IHttpPostMethod
	{
		private readonly HttpAuthenticationScheme[] authenticationSchemes;

		/// <summary>
		/// Converts a Excel document to Script
		/// </summary>
		/// <param name="AuthenticationSchemes">Authentication schemes.</param>
		public ExcelToScript(params HttpAuthenticationScheme[] AuthenticationSchemes)
			: base("/MicrosoftInterop/ExcelToScript")
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

			if (!(Decoded.Decoded is SpreadsheetDocument Doc))
			{
				await Response.SendResponse(new BadRequestException("Content not an Excel document (.xlsx)."));
				return;
			}

			StringBuilder Script = new StringBuilder();
			ExcelUtilities.ExtractAsScript(Doc, string.Empty, Script, true, out _);

			byte[] Data = WordToMarkdownConverter.Utf8WithBOM.GetBytes(Script.ToString());

			Response.ContentType = "text/plain; charset=utf-8";
			await Response.Write(true, Data);
		}
	}
}
