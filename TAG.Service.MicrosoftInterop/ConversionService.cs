using System;
using System.Threading.Tasks;
using Waher.IoTGateway;

namespace TAG.Service.MicrosoftInterop
{
	/// <summary>
	/// Conversion service for Microsoft technologies.
	/// </summary>
	public class ConversionService : IConfigurableModule
	{
		public ConversionService()
		{
		}

		/// <summary>
		/// Starts the service.
		/// </summary>
		public Task Start()
		{
			return Task.CompletedTask;
		}

		/// <summary>
		/// Stops the service.
		/// </summary>
		public Task Stop()
		{
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
