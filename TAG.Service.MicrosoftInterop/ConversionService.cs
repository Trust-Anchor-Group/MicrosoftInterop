using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Waher.IoTGateway;
using Waher.IoTGateway.Setup;
using Waher.Networking.HTTP.Authentication;
using Waher.Networking.HTTP;
using Waher.Networking;
using Waher.Runtime.Inventory;
using Waher.Security.JWT;
using Waher.Security.Users;
using TAG.Service.MicrosoftInterop.WebServices;
using Waher.Things.Http;

namespace TAG.Service.MicrosoftInterop
{
	/// <summary>
	/// Conversion service for Microsoft technologies.
	/// </summary>
	[ModuleDependency(typeof(HttpModule))]
	public class ConversionService : IConfigurableModule
	{
		private WordToMarkdown wordToMarkdown;
		private ExcelToScript excelToScript;
		private AppendingMarkdownLabMd appendingMarkdownLabMd;
		private AppendingMarkdownLabJs appendingMarkdownLabJs;
		private AppendingMarkdownLabCss appendingMarkdownLabCss;
		private AppendingPromptMd appendingPromptMd;
		private AppendingPromptJs appendingPromptJs;

		public ConversionService()
		{
		}

		/// <summary>
		/// Starts the service.
		/// </summary>
		public Task Start()
		{
			HttpAuthenticationScheme[] Schemes = HttpModule.GetAuthenticationSchemes();

			this.wordToMarkdown = new WordToMarkdown(Schemes);
			Gateway.HttpServer?.Register(this.wordToMarkdown);

			this.excelToScript = new ExcelToScript(Schemes);
			Gateway.HttpServer?.Register(this.excelToScript);

			Schemes = HttpModule.GetAuthenticationSchemes("Admin.Lab.Markdown", "Admin.Lab.Script");

			this.appendingMarkdownLabMd = new AppendingMarkdownLabMd(Schemes);
			Gateway.HttpServer?.Register(this.appendingMarkdownLabMd);

			this.appendingMarkdownLabJs = new AppendingMarkdownLabJs(Schemes);
			Gateway.HttpServer?.Register(this.appendingMarkdownLabJs);

			this.appendingMarkdownLabCss = new AppendingMarkdownLabCss(Schemes);
			Gateway.HttpServer?.Register(this.appendingMarkdownLabCss);

			Schemes = HttpModule.GetAuthenticationSchemes("Admin.Lab.Script");

			this.appendingPromptMd = new AppendingPromptMd(Schemes);
			Gateway.HttpServer?.Register(this.appendingPromptMd);

			this.appendingPromptJs = new AppendingPromptJs(Schemes);
			Gateway.HttpServer?.Register(this.appendingPromptJs);

			return Task.CompletedTask;
		}

		/// <summary>
		/// Stops the service.
		/// </summary>
		public Task Stop()
		{
			if (!(this.wordToMarkdown is null))
			{
				Gateway.HttpServer?.Unregister(this.wordToMarkdown);
				this.wordToMarkdown = null;
			}

			if (!(this.excelToScript is null))
			{
				Gateway.HttpServer?.Unregister(this.excelToScript);
				this.excelToScript = null;
			}

			if (!(this.appendingMarkdownLabMd is null))
			{
				Gateway.HttpServer?.Unregister(this.appendingMarkdownLabMd);
				this.appendingMarkdownLabMd = null;
			}

			if (!(this.appendingMarkdownLabJs is null))
			{
				Gateway.HttpServer?.Unregister(this.appendingMarkdownLabJs);
				this.appendingMarkdownLabJs = null;
			}

			if (!(this.appendingMarkdownLabCss is null))
			{
				Gateway.HttpServer?.Unregister(this.appendingMarkdownLabCss);
				this.appendingMarkdownLabCss = null;
			}

			if (!(this.appendingPromptMd is null))
			{
				Gateway.HttpServer?.Unregister(this.appendingPromptMd);
				this.appendingPromptMd = null;
			}

			if (!(this.appendingPromptJs is null))
			{
				Gateway.HttpServer?.Unregister(this.appendingPromptJs);
				this.appendingPromptJs = null;
			}

			return Task.CompletedTask;
		}

		/// <summary>
		/// Gets an array of pages used to configure the service.
		/// </summary>
		/// <returns>Configurable pages.</returns>
		public Task<IConfigurablePage[]> GetConfigurablePages()
		{
			return Task.FromResult(Array.Empty<IConfigurablePage>());
		}

	}
}
