using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Text;
using Waher.Runtime.Text;
using Waher.Script;

namespace TAG.Content.Microsoft.Test
{
	[TestClass]
	public class ExcelTests
	{
		private static string? inputFolder;
		private static string? outputFolder;
		private static string? expectedFolder;

		[ClassInitialize]
		public static Task ClassInitialize(TestContext _)
		{
			inputFolder = Path.Combine(Environment.CurrentDirectory, "Spreadsheets");
			outputFolder = Path.Combine(Environment.CurrentDirectory, "Output");
			expectedFolder = Path.Combine(Environment.CurrentDirectory, "Expected", "Script");

			if (!Directory.Exists(outputFolder))
				Directory.CreateDirectory(outputFolder);

			return Task.CompletedTask;
		}

		[DataTestMethod]
		[DataRow("SimpleSheet")]
		public async Task Convert_To_Script(string FileName)
		{
			Assert.IsNotNull(inputFolder);
			Assert.IsNotNull(outputFolder);
			Assert.IsNotNull(expectedFolder);

			string InputFileName = Path.Combine(inputFolder, FileName + ".xlsx");
			string OutputFileName = Path.Combine(outputFolder, FileName + ".script");
			string OutputFileName2 = Path.Combine(outputFolder, FileName + "2.script");
			string ExpectedFileName = Path.Combine(expectedFolder, FileName + ".script");

			ExcelUtilities.ConvertExcelToScript(InputFileName, OutputFileName, true);
			ExcelUtilities.ConvertExcelToScript(InputFileName, OutputFileName2, false);

			string Output = File.ReadAllText(OutputFileName);
			string Output2 = File.ReadAllText(OutputFileName2);

			Expression Exp1 = new(Output);
			Expression Exp2 = new(Output2);

			Variables Variables1 = new();
			object Result1 = await Exp1.EvaluateAsync(Variables1);
			string s1 = Expression.ToString(Result1);

			Variables Variables2 = new();
			object Result2 = await Exp2.EvaluateAsync(Variables2);
			string s2 = Expression.ToString(Result2);

			Assert.AreEqual(s1, s2);

			if (File.Exists(ExpectedFileName))
			{
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

		/* TODO:
		 * - Multiple sheets
		 * - References
		 * - Data types: Compare string "10" with number 10.
		 * - Normalized strings
		 * - Sparse matrices
		 * - Leading empty rows/columns
		 * - Multi-character columns & rows
		 */
	}
}