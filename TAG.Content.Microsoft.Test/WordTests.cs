using Microsoft.VisualStudio.TestTools.UnitTesting;
using Waher.Events;
using Waher.Events.Console;
using Waher.Runtime.Inventory.Loader;

namespace TAG.Content.Microsoft.Test
{
	[TestClass]
	public class WordTests
	{
		private static string? inputFolder;
		private static string? outputFolder;

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

			if (!Directory.Exists(outputFolder))
				Directory.CreateDirectory(outputFolder);

			return Task.CompletedTask;
		}

		[DataTestMethod]
		[DataRow("SimpleText")]
		public void Test_01_Convert_To_Markdown(string FileName)
		{
			Assert.IsNotNull(inputFolder);
			Assert.IsNotNull(outputFolder);

			string InputFileName = Path.Combine(inputFolder, FileName + ".docx");
			string OutputFileName = Path.Combine(outputFolder, FileName + ".md");

			WordUtilities.ConvertWordToMarkdown(InputFileName, OutputFileName);
		}

		/* Sections
		 * Columns
		 * Paragraph justification
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
		 */
	}
}