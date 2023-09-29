using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Text;
using Waher.Events;
using Waher.Events.Console;
using Waher.Runtime.Inventory.Loader;
using Waher.Runtime.Text;

namespace TAG.Content.Microsoft.Test
{
	[TestClass]
	public class WordTests
	{
		private static string? inputFolder;
		private static string? outputFolder;
		private static string? expectedFolder;

		[AssemblyInitialize]
		public static Task AssemblyInitialize(TestContext _)
		{
			// Create inventory of available classes.
			TypesLoader.Initialize();

			Log.Register(new ConsoleEventSink());

			return Task.CompletedTask;
		}

		[ClassInitialize]
		public static Task ClassInitialize(TestContext _)
		{
			inputFolder = Path.Combine(Environment.CurrentDirectory, "Documents");
			outputFolder = Path.Combine(Environment.CurrentDirectory, "Output");
			expectedFolder = Path.Combine(Environment.CurrentDirectory, "Expected", "Markdown");

			if (!Directory.Exists(outputFolder))
				Directory.CreateDirectory(outputFolder);

			return Task.CompletedTask;
		}

		[DataTestMethod]
		[DataRow("SimpleText")]
		[DataRow("Paragraphs")]
		[DataRow("Sections")]
		[DataRow("Tables")]
		[DataRow("Lists")]
		[DataRow("TableOfContents")]
		[DataRow("Figures")]
		[DataRow("Frames")]
		[DataRow("Notes")]
		[DataRow("Fields")]
		[DataRow("Fields")]
		[DataRow("MultiParagraphList")]
		public void Convert_To_Markdown(string FileName)
		{
			Assert.IsNotNull(inputFolder);
			Assert.IsNotNull(outputFolder);
			Assert.IsNotNull(expectedFolder);

			string InputFileName = Path.Combine(inputFolder, FileName + ".docx");
			string OutputFileName = Path.Combine(outputFolder, FileName + ".md");
			string ExpectedFileName = Path.Combine(expectedFolder, FileName + ".md");

			WordUtilities.ConvertWordToMarkdown(InputFileName, OutputFileName);

			if (File.Exists(ExpectedFileName))
			{
				string Output = File.ReadAllText(OutputFileName);
				string Expected = File.ReadAllText(ExpectedFileName);

				if (Expected != Output)
				{
					StringBuilder Error = new();

					Error.AppendLine("Output not as expected.");
					Error.AppendLine();

					foreach (Step<string> Change in Difference.AnalyzeRows(Output, Expected).Steps)
					{
						switch (Change.Operation)
						{
							case EditOperation.Insert:
								foreach (string Row in Change.Symbols)
								{
									Error.Append("+ ");
									Error.AppendLine(Row);
								}
								break;

							case EditOperation.Delete:
								foreach (string Row in Change.Symbols)
								{
									Error.Append("- ");
									Error.AppendLine(Row);
								}
								break;
						}
					}

					Assert.Fail(Error.ToString());
				}
			}
		}
	}
}