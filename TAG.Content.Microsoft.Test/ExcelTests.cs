using System.Diagnostics;
using System.Text;
using Waher.Content.Semantic;
using Waher.Runtime.Text;
using Waher.Script;
using Waher.Script.Abstraction.Elements;

namespace TAG.Content.Microsoft.Test
{
	[TestClass]
	public class ExcelTests
	{
		private static string? inputSpreadsheetsFolder;
		private static string? inputScriptFolder;
		private static string? outputFolder;
		private static string? expectedFolder;

		[ClassInitialize]
		public static Task ClassInitialize(TestContext _)
		{
			inputSpreadsheetsFolder = Path.Combine(Environment.CurrentDirectory, "Spreadsheets");
			inputScriptFolder = Path.Combine(Environment.CurrentDirectory, "Script");
			outputFolder = Path.Combine(Environment.CurrentDirectory, "Output");
			expectedFolder = Path.Combine(Environment.CurrentDirectory, "Expected", "Script");

			if (!Directory.Exists(outputFolder))
				Directory.CreateDirectory(outputFolder);

			return Task.CompletedTask;
		}

		[TestMethod]
		[DataRow("SimpleSheet")]
		[DataRow("MultipleSheets")]
		[DataRow("SparseMatrix")]
		[DataRow("Diagram")]
		[DataRow("Image")]
		public async Task Convert_To_Script(string FileName)
		{
			Assert.IsNotNull(inputSpreadsheetsFolder);
			Assert.IsNotNull(outputFolder);
			Assert.IsNotNull(expectedFolder);

			string InputFileName = Path.Combine(inputSpreadsheetsFolder, FileName + ".xlsx");
			string OutputFileName = Path.Combine(outputFolder, FileName + ".script");
			string OutputFileName2 = Path.Combine(outputFolder, FileName + "2.script");
			string ExpectedFileName = Path.Combine(expectedFolder, FileName + ".script");

			ExcelUtilities.ConvertExcelToScript(InputFileName, OutputFileName, true);
			ExcelUtilities.ConvertExcelToScript(InputFileName, OutputFileName2, false);

			string Output = File.ReadAllText(OutputFileName);
			string Output2 = File.ReadAllText(OutputFileName2);

			Expression Exp1 = new(Output);
			Expression Exp2 = new(Output2);

			Variables Variables1 = [];
			object Result1 = await Exp1.EvaluateAsync(Variables1);
			string s1 = Expression.ToString(Result1);

			Variables Variables2 = [];
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

		[TestMethod]
		[DataRow("MultiplicationTable")]
		[DataRow("SineTable")]
		[DataRow("SparqlResult")]
		[DataRow("TableUnits")]
		public async Task Convert_To_Excel(string FileName)
		{
			Assert.IsNotNull(inputScriptFolder);
			Assert.IsNotNull(outputFolder);

			string InputFileName = Path.Combine(inputScriptFolder, FileName + ".script");
			string OutputFileName = Path.Combine(outputFolder, FileName + ".xlsx");

			string Script = File.ReadAllText(InputFileName);
			object Value = await Expression.EvalAsync(Script, []);
			if (Value is IMatrix Matrix)
				ExcelUtilities.ConvertMatrixToExcel(Matrix, OutputFileName, "Result");
			else if (Value is SparqlResultSet SparqlResultSet)
				ExcelUtilities.ConvertResultSetToExcel(SparqlResultSet, OutputFileName, "Result");
			else
				throw new Exception("Unsupported result.");


			Process.Start(new ProcessStartInfo()
			{
				FileName = OutputFileName,
				UseShellExecute = true
			});
		}

	}
}