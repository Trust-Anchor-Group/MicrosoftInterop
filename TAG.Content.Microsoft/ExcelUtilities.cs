using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.ExcelAc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Waher.Content;
using Waher.Events;
using Waher.Script;

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
		/// <param name="Indentation">If Indentation is to be used.</param>
		public static void ConvertExcelToScript(string ExcelFileName, string ScriptFileName,
			bool Indentation)
		{
			string Script = ExtractAsScript(ExcelFileName, Indentation);
			File.WriteAllText(ScriptFileName, Script, WordUtilities.utf8BomEncoding);
		}


		/// <summary>
		/// Extracts the contents of a Excel file to Script.
		/// </summary>
		/// <param name="ExcelFileName">File name of Excel document.</param>
		/// <param name="Indentation">If Indentation is to be used.</param>
		/// <returns>Script</returns>
		public static string ExtractAsScript(string ExcelFileName, bool Indentation)
		{
			return ExtractAsScript(ExcelFileName, Indentation, out _);
		}

		/// <summary>
		/// Extracts the contents of a Excel file to Script.
		/// </summary>
		/// <param name="ExcelFileName">File name of Excel document.</param>
		/// <param name="Indentation">If Indentation is to be used.</param>
		/// <param name="Language">Language of document.</param>
		/// <returns>Script</returns>
		public static string ExtractAsScript(string ExcelFileName, bool Indentation,
			out string Language)
		{
			StringBuilder Script = new StringBuilder();
			ExtractAsScript(ExcelFileName, Script, Indentation, out Language);
			return Script.ToString();
		}

		/// <summary>
		/// Extracts the contents of a Excel file to Script.
		/// </summary>
		/// <param name="ExcelFileName">File name of Excel document.</param>
		/// <param name="Script">Script will be output here.</param>
		/// <param name="Indentation">If Indentation is to be used.</param>
		public static void ExtractAsScript(string ExcelFileName, StringBuilder Script,
			bool Indentation)
		{
			ExtractAsScript(ExcelFileName, Script, Indentation, out _);
		}

		/// <summary>
		/// Extracts the contents of a Excel file to Script.
		/// </summary>
		/// <param name="ExcelFileName">File name of Excel document.</param>
		/// <param name="Script">Script will be output here.</param>
		/// <param name="Indentation">If Indentation is to be used.</param>
		/// <param name="Language">Language of document.</param>
		public static void ExtractAsScript(string ExcelFileName, StringBuilder Script,
			bool Indentation, out string Language)
		{
			using (SpreadsheetDocument Doc = SpreadsheetDocument.Open(ExcelFileName, false))
			{
				ExtractAsScript(Doc, ExcelFileName, Script, Indentation, out Language);
			}
		}

		/// <summary>
		/// Extracts the contents of a Excel file to Script.
		/// </summary>
		/// <param name="Doc">Document to convert</param>
		/// <param name="Indentation">If Indentation is to be used.</param>
		/// <returns>Script</returns>
		public static string ExtractAsScript(SpreadsheetDocument Doc, bool Indentation)
		{
			return ExtractAsScript(Doc, Indentation, out _);
		}

		/// <summary>
		/// Extracts the contents of a Excel file to Script.
		/// </summary>
		/// <param name="Doc">Document to convert</param>
		/// <param name="Indentation">If Indentation is to be used.</param>
		/// <param name="Language">Language of document.</param>
		/// <returns>Script</returns>
		public static string ExtractAsScript(SpreadsheetDocument Doc, bool Indentation,
			out string Language)
		{
			return ExtractAsScript(Doc, string.Empty, Indentation, out Language);
		}

		/// <summary>
		/// Extracts the contents of a Excel file to Script.
		/// </summary>
		/// <param name="Doc">Document to convert</param>
		/// <param name="ExcelFileName">File name of Excel document.</param>
		/// <param name="Indentation">If Indentation is to be used.</param>
		/// <param name="Language">Language of document.</param>
		/// <returns>Script</returns>
		public static string ExtractAsScript(SpreadsheetDocument Doc, string ExcelFileName,
			bool Indentation, out string Language)
		{
			StringBuilder Script = new StringBuilder();
			ExtractAsScript(Doc, ExcelFileName, Script, Indentation, out Language);
			return Script.ToString();
		}

		/// <summary>
		/// Extracts the contents of a Excel file to Script.
		/// </summary>
		/// <param name="Doc">Document to convert</param>
		/// <param name="ExcelFileName">File name of Excel document.</param>
		/// <param name="Script">Script will be output here.</param>
		/// <param name="Indentation">If Indentation is to be used.</param>
		/// <param name="Language">Language of document.</param>
		public static void ExtractAsScript(SpreadsheetDocument Doc, string ExcelFileName, StringBuilder Script,
			bool Indentation, out string Language)
		{
			RenderingState State = new RenderingState()
			{
				Doc = Doc,
				Script = Script,
				Indentation = Indentation,
				NrTabs = 0,
			};

			Language = Doc.PackageProperties.Language;

			ExportAsScript(Doc.WorkbookPart.Workbook.Elements(), Script, State);

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

		private const string Excel2006Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
		private const string Compatibility2006Namespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";
		private const string Revision2014Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
		private const string Excel2010AbsoluteClassNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac";
		private const string Excel2010MainNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
		private const string Excel2018CalcFeaturesNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures";

		private static bool ExportAsScript(IEnumerable<OpenXmlElement> Elements,
			StringBuilder Script, RenderingState State)
		{
			bool HasText = false;

			foreach (OpenXmlElement Element in Elements)
			{
				if (ExportAsScript(Element, Script, State))
					HasText = true;
			}

			return HasText;
		}

		private static bool ExportAsScript(OpenXmlElement Element, StringBuilder Script,
			RenderingState State)
		{
			bool HasText = false;

			switch (Element.NamespaceUri)
			{
				case Excel2006Namespace:
					switch (Element.LocalName)
					{
						case "sheet":
							if (Element is Sheet Sheet)
							{
								State.Left = 0;
								State.Right = 0;
								State.Top = 0;
								State.Bottom = 0;
								State.Cells = null;

								Script.Append('{');
								State.NrTabs++;
								State.NewLine();

								Script.Append("'Name':'");
								if (Sheet.Name.HasValue)
									Script.Append(JSON.Encode(Sheet.Name.Value));

								Script.Append("',");
								State.NewLine();

								OpenXmlPart Part = State.Doc.WorkbookPart.GetPartById(Sheet.Id);
								HasText = ExportAsScript(Part.RootElement, Script, State);

								Script.Append("'Left':");
								Script.Append(State.Left.ToString());
								Script.Append(',');
								State.NewLine();

								Script.Append("'Top':");
								Script.Append(State.Top.ToString());
								Script.Append(',');
								State.NewLine();

								Script.Append("'Right':");
								Script.Append(State.Right.ToString());
								Script.Append(',');
								State.NewLine();

								Script.Append("'Bottom':");
								Script.Append(State.Bottom.ToString());
								Script.Append(',');
								State.NewLine();

								Script.Append("'Data':");
								State.NrTabs++;

								if (State.Cells is null)
									Script.Append("null");
								else
								{
									int x, y;
									int w = State.Width;
									int h = State.Height;
									string s;

									State.NewLine();
									Script.Append('[');
									State.NrSpaces++;

									for (y = 0; y < h; y++)
									{
										if (y > 0)
										{
											Script.Append(',');
											State.NewLine();
										}

										Script.Append('[');

										for (x = 0; x < w; x++)
										{
											if (x > 0)
												Script.Append(',');

											s = State.Cells[x, y];
											if (!string.IsNullOrEmpty(s))
												Script.Append(s);
										}

										Script.Append(']');
									}

									State.NrSpaces--;
									Script.Append(']');
								}

								State.NrTabs -= 2;
								State.NewLine();
								Script.Append('}');
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "worksheet":
							if (Element is Worksheet Worksheet)
								HasText = ExportAsScript(Worksheet.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "workbook":
							if (Element is Workbook Workbook)
								HasText = ExportAsScript(Workbook.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "fileVersion":
							if (Element is FileVersion FileVersion)
								HasText = ExportAsScript(FileVersion.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "workbookPr":
							if (Element is WorkbookProperties WorkbookProperties)
								HasText = ExportAsScript(WorkbookProperties.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "bookViews":
							if (Element is BookViews BookViews)
								HasText = ExportAsScript(BookViews.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "sheets":
							if (Element is Sheets Sheets)
							{
								bool First = true;

								Script.Append('{');
								State.NrTabs++;
								State.NewLine();

								foreach (Sheet Sheet2 in Sheets.Elements<Sheet>())
								{
									if (First)
										First = false;
									else
									{
										Script.Append(',');
										State.NewLine();
									}

									Script.Append('\'');
									if (Sheet2.Name.HasValue)
										Script.Append(JSON.Encode(Sheet2.Name.Value));

									Script.Append("':");
									State.NrTabs++;
									State.NewLine();
									ExportAsScript(Sheet2, Script, State);
									State.NrTabs--;
								}

								if (!First)
									State.NewLine();

								Script.Append('}');
								HasText = true;
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "calcPr":
							if (Element is CalculationProperties CalculationProperties)
								HasText = ExportAsScript(CalculationProperties.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "extLst":
							if (Element is WorkbookExtensionList WorkbookExtensionList)
								HasText = ExportAsScript(WorkbookExtensionList.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "dimension":
							if (Element is SheetDimension SheetDimension)
							{
								if (SheetDimension.Reference.HasValue &&
									ParseBoxReference(SheetDimension.Reference.Value,
									out int Left, out int Top, out int Right, out int Bottom))
								{
									State.Left = Left;
									State.Top = Top;
									State.Right = Right;
									State.Bottom = Bottom;
									State.Cells = new string[State.Width, State.Height];
								}
								else
								{
									State.Left = 0;
									State.Top = 0;
									State.Right = 0;
									State.Bottom = 0;
									State.Cells = null;
								}

								HasText = ExportAsScript(SheetDimension.Elements(), Script, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "sheetViews":
							if (Element is SheetViews SheetViews)
								HasText = ExportAsScript(SheetViews.Elements(), Script, State);
							else if (Element is ChartSheetViews ChartSheetViews)
								HasText = ExportAsScript(ChartSheetViews.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "sheetFormatPr":
							if (Element is SheetFormatProperties SheetFormatProperties)
								HasText = ExportAsScript(SheetFormatProperties.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "cols":
							if (Element is Columns Columns)
								HasText = ExportAsScript(Columns.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "sheetData":
							if (Element is SheetData SheetData)
								HasText = ExportAsScript(SheetData.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "pageMargins":
							if (Element is PageMargins PageMargins)
								HasText = ExportAsScript(PageMargins.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "pageSetup":
							if (Element is PageSetup PageSetup)
								HasText = ExportAsScript(PageSetup.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "sheetView":
							if (Element is SheetView SheetView)
								HasText = ExportAsScript(SheetView.Elements(), Script, State);
							else if (Element is ChartSheetView ChartSheetView)
								HasText = ExportAsScript(ChartSheetView.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "col":
							if (Element is Column Column)
								HasText = ExportAsScript(Column.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "row":
							if (Element is Row Row)
								HasText = ExportAsScript(Row.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "selection":
							if (Element is DocumentFormat.OpenXml.Spreadsheet.Selection Selection)
								HasText = ExportAsScript(Selection.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "c":
							if (Element is Cell Cell)
							{
								if (Cell.CellReference.HasValue &&
									ParseCellReference(Cell.CellReference.Value, out int X, out int Y))
								{
									State.X = X;
									State.Y = Y;
									State.Type = Cell.DataType?.Value;
								}
								else
								{
									State.X = -1;
									State.Y = -1;
									State.Type = null;
								}

								HasText = ExportAsScript(Cell.Elements(), Script, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "v":
							if (Element is CellValue CellValue)
							{
								if (State.X >= State.Left && State.Y >= State.Top &&
									State.X <= State.Right && State.Y <= State.Bottom &&
									!(State.Cells is null))
								{
									string s = CellValue.InnerText;

									if (State.Type.HasValue)
									{
										CellValues Value = State.Type.Value;

										if (Value == CellValues.SharedString)
										{
											if (int.TryParse(s, out int i) &&
												State.TryGetSharedString(i, out string s2))
											{
												State.Cells[State.X - State.Left, State.Y - State.Top] = Expression.ToString(s2);
											}
											else
												State.Cells[State.X - State.Left, State.Y - State.Top] = Expression.ToString(s);
										}
										else if (Value == CellValues.Date)
										{
											if (DateTime.TryParse(s, out DateTime TP))
												State.Cells[State.X - State.Left, State.Y - State.Top] = Expression.ToString(TP);
											else
												State.Cells[State.X - State.Left, State.Y - State.Top] = Expression.ToString(s);
										}
										else if (Value == CellValues.Boolean)
										{
											if (CommonTypes.TryParse(s, out bool b))
												State.Cells[State.X - State.Left, State.Y - State.Top] = Expression.ToString(b);
											else
												State.Cells[State.X - State.Left, State.Y - State.Top] = Expression.ToString(s);
										}
										else if (Value == CellValues.Number)
										{
											if (CommonTypes.TryParse(s, out double d))
												State.Cells[State.X - State.Left, State.Y - State.Top] = Expression.ToString(d);
											else
												State.Cells[State.X - State.Left, State.Y - State.Top] = Expression.ToString(s);
										}
										else if (Value == CellValues.Error || Value == CellValues.String || Value == CellValues.InlineString)
											State.Cells[State.X - State.Left, State.Y - State.Top] = Expression.ToString(s);
									}
									else
									{
										if (CommonTypes.TryParse(s, out double d))
											State.Cells[State.X - State.Left, State.Y - State.Top] = Expression.ToString(d);
										else if (CommonTypes.TryParse(s, out bool b))
											State.Cells[State.X - State.Left, State.Y - State.Top] = Expression.ToString(b);
										else
											State.Cells[State.X - State.Left, State.Y - State.Top] = Expression.ToString(s);
									}
								}

								HasText = ExportAsScript(CellValue.Elements(), Script, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "f":
							if (Element is CellFormula CellFormula)
								HasText = ExportAsScript(CellFormula.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "workbookView":
							if (Element is WorkbookView WorkbookView)
								HasText = ExportAsScript(WorkbookView.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "ext":
							if (Element is WorkbookExtension WorkbookExtension)
								HasText = ExportAsScript(WorkbookExtension.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "chartsheet":
							if (Element is Chartsheet Chartsheet)
								HasText = ExportAsScript(Chartsheet.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "drawing":
							if (Element is Drawing Drawing)
								HasText = ExportAsScript(Drawing.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "sheetPr":
							if (Element is ChartSheetProperties ChartSheetProperties)
								HasText = ExportAsScript(ChartSheetProperties.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						default:
							State.UnrecognizedElement(Element);
							break;
					}
					break;

				case Compatibility2006Namespace:
					switch (Element.LocalName)
					{
						case "AlternateContent":
							if (Element is AlternateContent AlternateContent)
								HasText = ExportAsScript(AlternateContent.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "Choice":
							if (Element is AlternateContentChoice AlternateContentChoice)
								HasText = ExportAsScript(AlternateContentChoice.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						default:
							State.UnrecognizedElement(Element);
							break;
					}
					break;

				case Excel2010AbsoluteClassNamespace:
					switch (Element.LocalName)
					{
						case "absPath":
							if (Element is AbsolutePath AbsolutePath)
								HasText = ExportAsScript(AbsolutePath.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						default:
							State.UnrecognizedElement(Element);
							break;
					}
					break;

				case Excel2010MainNamespace:
					switch (Element.LocalName)
					{
						case "workbookPr":
							if (Element is DocumentFormat.OpenXml.Office2013.Excel.WorkbookProperties WorkbookProperties)
								HasText = ExportAsScript(WorkbookProperties.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						default:
							State.UnrecognizedElement(Element);
							break;
					}
					break;

				case Excel2018CalcFeaturesNamespace:
					switch (Element.LocalName)
					{
						case "calcFeatures":
							if (Element is OpenXmlUnknownElement OpenXmlUnknownElement)
								HasText = ExportAsScript(OpenXmlUnknownElement.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "feature":
							if (Element is OpenXmlUnknownElement Feature)
								HasText = ExportAsScript(Feature.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						default:
							State.UnrecognizedElement(Element);
							break;
					}
					break;

				case Revision2014Namespace:
					switch (Element.LocalName)
					{
						case "revisionPtr":
							if (Element is OpenXmlUnknownElement OpenXmlUnknownElement)
								HasText = ExportAsScript(OpenXmlUnknownElement.Elements(), Script, State);
							else
								State.UnrecognizedElement(Element);
							break;

						default:
							State.UnrecognizedElement(Element);
							break;
					}
					break;

				default:
					State.UnrecognizedElement(Element);
					break;
			}

			return HasText;
		}

		private static bool ParseBoxReference(string Ref, out int Left, out int Top,
			out int Right, out int Bottom)
		{
			Left = Top = Right = Bottom = 0;

			int i = Ref.IndexOf(':');
			if (i < 0)
				return false;

			return ParseCellReference(Ref.Substring(0, i), out Left, out Top) &&
				ParseCellReference(Ref.Substring(i + 1), out Right, out Bottom);
		}

		private static bool ParseCellReference(string Ref, out int Column, out int Row)
		{
			bool HasColumn = false;
			bool HasRow = false;
			char ch2;
			int i;

			Column = 0;
			Row = 0;

			foreach (char ch in Ref)
			{
				ch2 = char.ToUpper(ch);

				if (ch2 >= 'A' && ch2 <= 'Z')
				{
					i = Column;

					Column *= 26;
					Column += ch2 - 'A' + 1;

					if (Column < i)
						return false;

					HasColumn = true;
				}
				else if (ch2 >= '0' && ch2 <= '9')
				{
					i = Row;

					Row *= 10;
					Row += ch2 - '0';

					if (Row < i)
						return false;

					HasRow = true;
				}
				else
					return false;
			}

			return HasColumn && HasRow;
		}

		private class RenderingState
		{
			public SpreadsheetDocument Doc;
			public StringBuilder Script;
			public Dictionary<string, Dictionary<string, int>> Unrecognized = null;
			public Dictionary<string, int> LanguageCounts = new Dictionary<string, int>();
			public List<string> SharedStrings = null;
			public bool Indentation;
			public string[,] Cells;
			public int NrTabs;
			public int NrSpaces;
			public int Left;
			public int Top;
			public int Right;
			public int Bottom;
			public int X;
			public int Y;
			public CellValues? Type;

			public int Width => this.Right - this.Left + 1;
			public int Height => this.Bottom - this.Top + 1;

			public void NewLine()
			{
				if (this.Indentation)
				{
					this.Script.AppendLine();

					int i;

					for (i = 0; i < this.NrTabs; i++)
						this.Script.Append('\t');

					for (i = 0; i < this.NrSpaces; i++)
						this.Script.Append(' ');
				}
			}

			public void UnrecognizedElement(OpenXmlElement Element)
			{
				if (this.Unrecognized is null)
					this.Unrecognized = new Dictionary<string, Dictionary<string, int>>();

				string Key = Element.LocalName;

				if (Element.NamespaceUri != Excel2006Namespace)
					Key = Element.NamespaceUri + "#" + Key;

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

			public bool TryGetSharedString(int Index, out string Value)
			{
				if (this.SharedStrings is null)
				{
					this.SharedStrings = new List<string>();

					if (!(this.Doc.WorkbookPart.SharedStringTablePart?.RootElement is null))
					{
						foreach (OpenXmlElement E in this.Doc.WorkbookPart.SharedStringTablePart.RootElement.Elements())
						{
							if (E is SharedStringItem SharedString)
								this.SharedStrings.Add(SharedString.InnerText);
						}
					}
				}

				if (Index < 0 || Index >= this.SharedStrings.Count)
				{
					Value = null;
					return false;
				}
				else
				{
					Value = this.SharedStrings[Index];
					return true;
				}
			}
		}
	}
}
