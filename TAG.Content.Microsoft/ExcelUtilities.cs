using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Waher.Content.Markdown.Functions;
using Waher.Content.Xml;
using Waher.Events;

namespace TAG.Content.Microsoft
{
	/// <summary>
	/// Utilities for interoperation with Microsoft Office Excel documents.
	/// </summary>
	public static class ExcelUtilities
	{
		/// <summary>
		/// Converts an Excel document to scrpit.
		/// </summary>
		/// <param name="ExcelFileName">File name of Excel document.</param>
		/// <param name="ScriptFileName">File name of script file.</param>
		public static void ConvertExcelToScript(string ExcelFileName, string ScriptFileName)
		{
			string Script = ExtractAsScript(ExcelFileName);
			File.WriteAllText(ScriptFileName, Script, WordUtilities.utf8BomEncoding);
		}


		/// <summary>
		/// Extracts the contents of a Excel file to Script.
		/// </summary>
		/// <param name="ExcelFileName">File name of Excel document.</param>
		/// <returns>Script</returns>
		public static string ExtractAsScript(string ExcelFileName)
		{
			return ExtractAsScript(ExcelFileName, out _);
		}

		/// <summary>
		/// Extracts the contents of a Excel file to Script.
		/// </summary>
		/// <param name="ExcelFileName">File name of Excel document.</param>
		/// <param name="Language">Language of document.</param>
		/// <returns>Script</returns>
		public static string ExtractAsScript(string ExcelFileName, out string Language)
		{
			StringBuilder Script = new StringBuilder();
			ExtractAsScript(ExcelFileName, Script, out Language);
			return Script.ToString();
		}

		/// <summary>
		/// Extracts the contents of a Excel file to Script.
		/// </summary>
		/// <param name="ExcelFileName">File name of Excel document.</param>
		/// <param name="Script">Script will be output here.</param>
		public static void ExtractAsScript(string ExcelFileName, StringBuilder Script)
		{
			ExtractAsScript(ExcelFileName, Script, out _);
		}

		/// <summary>
		/// Extracts the contents of a Excel file to Script.
		/// </summary>
		/// <param name="ExcelFileName">File name of Excel document.</param>
		/// <param name="Script">Script will be output here.</param>
		/// <param name="Language">Language of document.</param>
		public static void ExtractAsScript(string ExcelFileName, StringBuilder Script,
			out string Language)
		{
			using (SpreadsheetDocument Doc = SpreadsheetDocument.Open(ExcelFileName, false))
			{
				ExtractAsScript(Doc, ExcelFileName, Script, out Language);
			}
		}

		/// <summary>
		/// Extracts the contents of a Excel file to Script.
		/// </summary>
		/// <param name="Doc">Document to convert</param>
		/// <returns>Script</returns>
		public static string ExtractAsScript(SpreadsheetDocument Doc)
		{
			return ExtractAsScript(Doc, out _);
		}

		/// <summary>
		/// Extracts the contents of a Excel file to Script.
		/// </summary>
		/// <param name="Doc">Document to convert</param>
		/// <param name="Language">Language of document.</param>
		/// <returns>Script</returns>
		public static string ExtractAsScript(SpreadsheetDocument Doc, out string Language)
		{
			return ExtractAsScript(Doc, string.Empty, out Language);
		}

		/// <summary>
		/// Extracts the contents of a Excel file to Script.
		/// </summary>
		/// <param name="Doc">Document to convert</param>
		/// <param name="ExcelFileName">File name of Excel document.</param>
		/// <param name="Language">Language of document.</param>
		/// <returns>Script</returns>
		public static string ExtractAsScript(SpreadsheetDocument Doc, string ExcelFileName, out string Language)
		{
			StringBuilder Script = new StringBuilder();
			ExtractAsScript(Doc, ExcelFileName, Script, out Language);
			return Script.ToString();
		}

		/// <summary>
		/// Extracts the contents of a Excel file to Script.
		/// </summary>
		/// <param name="Doc">Document to convert</param>
		/// <param name="ExcelFileName">File name of Excel document.</param>
		/// <param name="Script">Script will be output here.</param>
		/// <param name="Language">Language of document.</param>
		public static void ExtractAsScript(SpreadsheetDocument Doc, string ExcelFileName, StringBuilder Script,
			out string Language)
		{
			RenderingState State = new RenderingState()
			{
				Doc = Doc
			};

			Language = Doc.PackageProperties.Language;

			ExportAsScript(Doc.WorkbookPart.Workbook.Sheets.Elements(), Script, State);

			if (string.IsNullOrEmpty(Language))
			{
				int Best = 0;

				foreach (KeyValuePair<string, int> P in State.LanguageCounts)
				{
					if (P.Value > Best)
					{
						Best = P.Value;
						Language = P.Key;
					}
				}
			}

			if (!(Doc.WorkbookPart.Workbook.Sheets is null))
			{
				foreach (OpenXmlElement E in Doc.WorkbookPart.Workbook.Sheets)
				{
				}
			}

			if (!(State.Unrecognized is null))
			{
				StringBuilder Msg = new StringBuilder();

				Msg.AppendLine("Open XML-spreadsheet with unrecognized elements converted to Script.");
				Msg.AppendLine();
				Msg.AppendLine("| Element | Open XML Type | Occurrences |");
				Msg.AppendLine("|:--------|:--------------|------------:|");

				foreach (KeyValuePair<string, Dictionary<string, int>> P in State.Unrecognized)
				{
					foreach (KeyValuePair<string, int> P2 in P.Value)
					{
						Msg.Append("| `");
						Msg.Append(P.Key);
						Msg.Append("` | `");
						Msg.Append(P2.Key);
						Msg.Append("` | ");
						Msg.Append(P2.Value.ToString());
						Msg.AppendLine(" |");
					}
				}
#if DEBUG
				Msg.AppendLine();
				Msg.AppendLine("```");
				Msg.AppendLine(WordUtilities.CapLength(XML.PrettyXml(Doc.WorkbookPart.Workbook.OuterXml), 256 * 1024));
				Msg.AppendLine("```");
#endif
				Log.Warning(Msg.ToString(), ExcelFileName);
			}
		}

		private class RenderingState
		{
			public SpreadsheetDocument Doc;
			public Dictionary<string, Dictionary<string, int>> Unrecognized = null;
			public Dictionary<string, int> LanguageCounts = new Dictionary<string, int>();

			public void UnrecognizedElement(OpenXmlElement Element)
			{
				if (this.Unrecognized is null)
					this.Unrecognized = new Dictionary<string, Dictionary<string, int>>();

				string Key = Element.NamespaceUri + "#" + Element.LocalName;

				if (!this.Unrecognized.TryGetValue(Key, out Dictionary<string, int> ByType))
				{
					ByType = new Dictionary<string, int>();
					this.Unrecognized[Key] = ByType;
				}

				string s = Element.GetType().FullName;

				if (!ByType.TryGetValue(s, out int i))
					i = 1;
				else
					i++;

				ByType[s] = i;
			}
		}

	}
}
