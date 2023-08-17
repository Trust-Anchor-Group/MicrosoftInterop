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
			expectedFolder = Path.Combine(Environment.CurrentDirectory, "Expected");

			if (!Directory.Exists(outputFolder))
				Directory.CreateDirectory(outputFolder);

			return Task.CompletedTask;
		}

		[DataTestMethod]
		[DataRow("SimpleText")]
		[DataRow("Paragraphs")]
		[DataRow("Sections")]
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

					foreach (Step<string> Change in Difference.AnalyzeRows(Expected, Output).Steps)
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

		/* 
		 * Code block
		 * Footnotes
		 * Tables
		 * Figures (Picture, Drawing)
		 * Table of contents
		 * Fields
		 * Horizontal Separator
		 * Bullet-lists
		 * Numbered lists
		 * Mixed lists
		 * Table column alignments.
		 * Table column spans.
		 * Table with multiple header rows
		 * Table captions
		 * Figure captions
		 * Cross-paragraph styles
		 * Cross-table-cell styles
		 * Cross-table-row styles
		 * Cross-table styles
		 * Footnotes
		 * Endnotes
		 * Detect document language
		 */
	}
}