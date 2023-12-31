﻿using System;
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
using Waher.IoTGateway.WebResources;

namespace TAG.Service.MicrosoftInterop
{
	/// <summary>
	/// Conversion service for Microsoft technologies.
	/// </summary>
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
			List<HttpAuthenticationScheme> Schemes = new List<HttpAuthenticationScheme>();
			bool RequireEncryption;
			int MinSecurityStrength;

			if (DomainConfiguration.Instance.UseEncryption && !string.IsNullOrEmpty(DomainConfiguration.Instance.Domain))
			{
				RequireEncryption = true;
				MinSecurityStrength = 128;
			}
			else
			{
				RequireEncryption = false;
				MinSecurityStrength = 0;
			}

			if (Types.TryGetModuleParameter("JWT", out object Obj) &&
				Obj is JwtFactory JwtFactory &&
				!JwtFactory.Disposed)
			{
				Schemes.Add(new JwtAuthentication(RequireEncryption, MinSecurityStrength, Gateway.Domain, null, JwtFactory));   // Any JWT token generated by the server will suffice. Does not have to point to a registered user.
			}

			if (!(Gateway.HttpServer is null) && Gateway.HttpServer.ClientCertificates != ClientCertificates.NotUsed)
				Schemes.Add(new MutualTlsAuthentication(Users.Source));

			Schemes.Add(new BasicAuthentication(RequireEncryption, MinSecurityStrength, Gateway.Domain, Users.Source));
			Schemes.Add(new DigestAuthentication(RequireEncryption, MinSecurityStrength, DigestAlgorithm.MD5, Gateway.Domain, Users.Source));
			Schemes.Add(new DigestAuthentication(RequireEncryption, MinSecurityStrength, DigestAlgorithm.SHA256, Gateway.Domain, Users.Source));
			Schemes.Add(new DigestAuthentication(RequireEncryption, MinSecurityStrength, DigestAlgorithm.SHA3_256, Gateway.Domain, Users.Source));
			Schemes.Add(new RequiredUserPrivileges(Gateway.HttpServer));

			this.wordToMarkdown = new WordToMarkdown(Schemes.ToArray());
			Gateway.HttpServer?.Register(this.wordToMarkdown);

			this.excelToScript = new ExcelToScript(Schemes.ToArray());
			Gateway.HttpServer?.Register(this.excelToScript);

			Schemes.Clear();
			Schemes.Add(new RequiredUserPrivileges("User", "/Login.md", Gateway.HttpServer, "Admin.Lab.Markdown", "Admin.Lab.Script"));

			this.appendingMarkdownLabMd = new AppendingMarkdownLabMd(Schemes.ToArray());
			Gateway.HttpServer?.Register(this.appendingMarkdownLabMd);

			this.appendingMarkdownLabJs = new AppendingMarkdownLabJs(Schemes.ToArray());
			Gateway.HttpServer?.Register(this.appendingMarkdownLabJs);

			this.appendingMarkdownLabCss = new AppendingMarkdownLabCss(Schemes.ToArray());
			Gateway.HttpServer?.Register(this.appendingMarkdownLabCss);

			Schemes.Clear();
			Schemes.Add(new RequiredUserPrivileges("User", "/Login.md", Gateway.HttpServer, "Admin.Lab.Script"));

			this.appendingPromptMd = new AppendingPromptMd(Schemes.ToArray());
			Gateway.HttpServer?.Register(this.appendingPromptMd);

			this.appendingPromptJs = new AppendingPromptJs(Schemes.ToArray());
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
