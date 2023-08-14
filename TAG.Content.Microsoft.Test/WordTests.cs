using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Text;
using Waher.Events.Console;
using Waher.Events;
using Waher.Persistence.Files;
using Waher.Persistence;
using Waher.Runtime.Inventory.Loader;
using Waher.Runtime.Inventory;
using Waher.Runtime.Settings;

namespace TAG.Content.Microsoft.Test
{
	[TestClass]
	public class WordTests
	{
		private static string? wordFileName;
		private static string? outputFolder;

		[AssemblyInitialize]
		public static async Task AssemblyInitialize(TestContext _)
		{
			// Create inventory of available classes.
			TypesLoader.Initialize();

			// Register console event log
			Log.Register(new ConsoleEventSink(true, true));

			// Instantiate local encrypted object database.
			FilesProvider DB = await FilesProvider.CreateAsync(Path.Combine(Directory.GetCurrentDirectory(), "Data"), "Default",
				8192, 10000, 8192, Encoding.UTF8, 10000, true, false);

			await DB.RepairIfInproperShutdown(string.Empty);

			Database.Register(DB);

			// Start embedded modules (database lifecycle)

			await Types.StartAllModules(60000);
		}

		[AssemblyCleanup]
		public static async Task AssemblyCleanup()
		{
			Log.Terminate();
			await Types.StopAllModules();
		}

		[ClassInitialize]
		public static async Task ClassInitialize(TestContext _)
		{
			//// Set to point to file used for tests.
			//await RuntimeSettings.SetAsync("WordFileName", @"");

			wordFileName = await RuntimeSettings.GetAsync("WordFileName", string.Empty);
			if (string.IsNullOrEmpty(wordFileName))
				throw new Exception("No Word file name has been configured.");

			outputFolder = Path.Combine(Environment.CurrentDirectory, "Output");
			if (!Directory.Exists(outputFolder))
				Directory.CreateDirectory(outputFolder);
		}

		[ClassCleanup]
		public static void ClassCleanup()
		{
			wordFileName = null;
		}

		[TestMethod]
		public void Test_01_Convert_To_PDF_Screen()
		{
			Assert.IsNotNull(outputFolder);
			string OutputFileName = Path.Combine(outputFolder, "Test_01.pdf");
			WordUtilities.ConvertWordToPdf(wordFileName, OutputFileName, false);
		}

		[TestMethod]
		public void Test_02_Convert_To_PDF_Print()
		{
			Assert.IsNotNull(outputFolder);
			string OutputFileName = Path.Combine(outputFolder, "Test_02.pdf");
			WordUtilities.ConvertWordToPdf(wordFileName, OutputFileName, true);
		}
	}
}