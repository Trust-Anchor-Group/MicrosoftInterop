using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.ExcelAc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Numerics;
using System.Text;
using Waher.Content;
using Waher.Content.Semantic;
using Waher.Content.Semantic.Model;
using Waher.Content.Xml;
using Waher.Events;
using Waher.Script;
using Waher.Script.Abstraction.Elements;
using Waher.Script.Objects;
using Waher.Script.Objects.Matrices;

namespace TAG.Content.Microsoft
{
	/// <summary>
	/// Utilities for interoperation with Microsoft Office Excel spreadsheets.
	/// </summary>
	public static class ExcelUtilities
	{
		#region Conversion of Excel spreadsheets to script

		/// <summary>
		/// Converts an Excel spreadsheet to script.
		/// </summary>
		/// <param name="ExcelFileName">File name of Excel spreadsheet.</param>
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
		/// <param name="ExcelFileName">File name of Excel spreadsheet.</param>
		/// <param name="Indentation">If Indentation is to be used.</param>
		/// <returns>Script</returns>
		public static string ExtractAsScript(string ExcelFileName, bool Indentation)
		{
			return ExtractAsScript(ExcelFileName, Indentation, out _);
		}

		/// <summary>
		/// Extracts the contents of a Excel file to Script.
		/// </summary>
		/// <param name="ExcelFileName">File name of Excel spreadsheet.</param>
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
		/// <param name="ExcelFileName">File name of Excel spreadsheet.</param>
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
		/// <param name="ExcelFileName">File name of Excel spreadsheet.</param>
		/// <param name="Script">Script will be output here.</param>
		/// <param name="Indentation">If Indentation is to be used.</param>
		/// <param name="Language">Language of document.</param>
		public static void ExtractAsScript(string ExcelFileName, StringBuilder Script,
			bool Indentation, out string Language)
		{
			using SpreadsheetDocument Doc = SpreadsheetDocument.Open(ExcelFileName, false);

			ExtractAsScript(Doc, ExcelFileName, Script, Indentation, out Language);
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
		/// <param name="ExcelFileName">File name of Excel spreadsheet.</param>
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
		/// <param name="ExcelFileName">File name of Excel spreadsheet.</param>
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
											else if (CommonTypes.TryParse(s, out double d))
												State.Cells[State.X - State.Left, State.Y - State.Top] = Expression.ToString(DateTime.FromOADate(d));
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

			return ParseCellReference(Ref[..i], out Left, out Top) &&
				ParseCellReference(Ref[(i + 1)..], out Right, out Bottom);
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
				this.Unrecognized ??= new Dictionary<string, Dictionary<string, int>>();

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

		#endregion

		#region Conversion of matrices to Excel spreadsheets

		private const double MaxDigitWidth = 7.1;

		/// <summary>
		/// Converts a matrix to an Excel spreadsheet.
		/// </summary>
		/// <param name="M">Matrix object instance.</param>
		/// <param name="FilePath">Path where Excel file will be stored.</param>
		/// <param name="SheetName">Name of sheet.</param>
		/// <returns>Spreadsheet document.</returns>
		public static void ConvertMatrixToExcel(IMatrix M, string FilePath, string SheetName)
		{
			if (M is null)
				throw new ArgumentNullException(nameof(M));

			using SpreadsheetDocument Document = SpreadsheetDocument.Create(FilePath, SpreadsheetDocumentType.Workbook);

			WorkbookPart WorkbookPart = Document.AddWorkbookPart();
			WorkbookPart.Workbook = new Workbook();

			WorksheetPart WorksheetPart = WorkbookPart.AddNewPart<WorksheetPart>();
			WorksheetPart.Worksheet = new Worksheet();

			WorkbookStylesPart StylesPart = WorkbookPart.AddNewPart<WorkbookStylesPart>();
			StylesPart.Stylesheet = CreateStylesheet();  // Define bold and normal styles
			StylesPart.Stylesheet.Save();

			Columns SheetColumns = new Columns();
			WorksheetPart.Worksheet.Append(SheetColumns);
			List<Column> SheetColumn = new List<Column>();

			SheetData SheetData = new SheetData();
			WorksheetPart.Worksheet.Append(SheetData);

			uint[] Increments = ColumnIncrements(M);
			uint ColumnTotal = ColumnCount(Increments);
			int Columns = M.Columns;
			int Rows = M.Rows;
			int Row, Column;
			uint x = 1, y = 1;  // Excel uses 1-based indices.
			string s;
			Row ExcelRow;
			Cell Cell;

			if (M is ObjectMatrix OM && OM.HasColumnNames)
			{
				int c = OM.ColumnNames.Length;
				ExcelRow = new Row()
				{
					RowIndex = y
				};

				for (Column = 0; Column < c; Column++)
				{
					s = OM.ColumnNames[Column];
					CheckColumnWidth(s, x, SheetColumns, SheetColumn);

					Cell = new Cell()
					{
						CellReference = GetCellReference(x, y),
						CellValue = new CellValue(s),
						DataType = CellValues.String,
						StyleIndex = 1  // Bold
					};

					ExcelRow.Append(Cell);
					x += Increments[Column];
				}

				x = 1;
				y++;

				SheetData.Append(ExcelRow);
			}

			for (Row = 0; Row < Rows; Row++)
			{
				ExcelRow = new Row()
				{
					RowIndex = y
				};

				for (Column = 0; Column < Columns; Column++)
				{
					object Value = M.GetElement(Column, Row)?.AssociatedObjectValue;

					if (!(Value is null))
					{
						s = Value.ToString();
						ExcelRow.Append(CreateCell(null, x, y, Value, ref s, out string Unit));
						CheckColumnWidth(s, x, SheetColumns, SheetColumn);

						if (!string.IsNullOrEmpty(Unit))
						{
							s = Unit;
							ExcelRow.Append(CreateCell(null, x + 1, y, Unit, ref s, out Unit));
							CheckColumnWidth(s, x + 1, SheetColumns, SheetColumn);
						}
					}

					x += Increments[Column];
				}

				x = 1;
				y++;

				SheetData.Append(ExcelRow);
			}

			Sheets Sheets = WorkbookPart.Workbook.AppendChild(new Sheets());
			Sheet Sheet = new Sheet()
			{
				Id = WorkbookPart.GetIdOfPart(WorksheetPart),
				SheetId = 1,
				Name = SheetName
			};
			Sheets.Append(Sheet);

			WorksheetPart.Worksheet.Save();
			WorkbookPart.Workbook.Save();
		}

		private static void CheckColumnWidth(string Content, uint x, Columns SheetColumns, List<Column> SheetColumn)
		{
			double Width = Math.Floor((Content.Length * MaxDigitWidth + 5) / MaxDigitWidth * 256) / 256;
			Column Column;

			if (x > SheetColumn.Count)
			{
				Column = new Column()
				{
					Min = x,
					Max = x,
					Width = Width,
					CustomWidth = true
				};
				SheetColumn.Add(Column);
				SheetColumns.Append(Column);
			}
			else
			{
				Column = SheetColumn[(int)(x - 1)];

				if (Width > Column.Width)
					Column.Width = Width;
			}
		}

		private static uint[] ColumnIncrements(IMatrix Result)
		{
			int x, Columns = Result.Columns;
			int y, Rows = Result.Rows;
			uint[] Increments = new uint[Columns];

			Array.Fill(Increments, (uint)1);

			for (y = 0; y < Rows; y++)
			{
				for (x = 0; x < Columns; x++)
				{
					if (Increments[x] == 1)
					{
						object Element = Result.GetElement(x, y)?.AssociatedObjectValue;

						if (Element is PhysicalQuantity)
							Increments[x] = 2;
					}
				}
			}

			return Increments;
		}

		private static uint ColumnCount(uint[] ColumnIncrements)
		{
			uint Result = 0;

			foreach (uint Increment in ColumnIncrements)
				Result += Increment;

			return Result;
		}

		private static Cell CreateCell(SparqlResultSet Result, uint Column, uint Row,
			object Value, ref string StringValue, out string Unit)
		{
			Cell Cell = new Cell()
			{
				CellReference = GetCellReference(Column, Row),
				StyleIndex = 0  // Normal
			};

			Unit = null;

			if (Value is ISemanticLiteral SemanticLiteral)
			{
				Value = SemanticLiteral.AssociatedObjectValue;
				StringValue = null;
			}

			if (Value is PhysicalQuantity Quantity)
			{
				Value = Quantity.Magnitude;
				Unit = Quantity.Unit.ToString();
				StringValue = null;
			}

			StringValue ??= Value.ToString();

			if (Value is string s)
			{
				if (s.IndexOf('\n') < 0)
				{
					Cell.CellValue = new CellValue(s);
					Cell.DataType = CellValues.String;
				}
				else
				{
					string[] Rows = s.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
					StringBuilder sb = new StringBuilder(Rows[0]);
					int MaxIndex = 0;
					int MaxLen = Rows[0].Length;
					int i, j, c = Rows.Length;

					for (i = 1; i < c; i++)
					{
						s = Rows[i];
						j = s.Length;
						if (j > MaxLen)
						{
							MaxLen = j;
							MaxIndex = i;
						}

						sb.Append("\r\n");
						sb.Append(s);
					}

					StringValue = Rows[MaxIndex];

					Text Text = new Text(sb.ToString())
					{
						Space = SpaceProcessingModeValues.Preserve
					};

					InlineString InlineString = new InlineString();
					InlineString.AppendChild(Text);

					Cell.InlineString = InlineString;
					Cell.DataType = CellValues.InlineString;
					Cell.StyleIndex = 2;    // Wrap text
				}
			}
			else if (Value is double d)
			{
				Cell.CellValue = new CellValue(d);
				Cell.DataType = CellValues.Number;

				if (StringValue.Length > 11)
					StringValue = StringValue[..11];
			}
			else if (Value is decimal dec)
			{
				Cell.CellValue = new CellValue(dec);
				Cell.DataType = CellValues.Number;

				if (StringValue.Length > 11)
					StringValue = StringValue[..11];
			}
			else if (Value is float f)
			{
				Cell.CellValue = new CellValue(f);
				Cell.DataType = CellValues.Number;

				if (StringValue.Length > 11)
					StringValue = StringValue[..11];
			}
			else if (Value is sbyte i8)
			{
				Cell.CellValue = new CellValue(i8);
				Cell.DataType = CellValues.Number;
			}
			else if (Value is short i16)
			{
				Cell.CellValue = new CellValue(i16);
				Cell.DataType = CellValues.Number;
			}
			else if (Value is int i32)
			{
				Cell.CellValue = new CellValue(i32);
				Cell.DataType = CellValues.Number;
			}
			else if (Value is long i64)
			{
				Cell.CellValue = new CellValue((decimal)i64);
				Cell.DataType = CellValues.Number;
			}
			else if (Value is byte ui8)
			{
				Cell.CellValue = new CellValue(ui8);
				Cell.DataType = CellValues.Number;
			}
			else if (Value is ushort ui16)
			{
				Cell.CellValue = new CellValue(ui16);
				Cell.DataType = CellValues.Number;
			}
			else if (Value is uint ui32)
			{
				Cell.CellValue = new CellValue((decimal)ui32);
				Cell.DataType = CellValues.Number;
			}
			else if (Value is ulong ui64)
			{
				Cell.CellValue = new CellValue((decimal)ui64);
				Cell.DataType = CellValues.Number;
			}
			else if (Value is BigInteger i)
			{
				Cell.CellValue = new CellValue(i.ToString());
				Cell.DataType = CellValues.Number;
			}
			else if (Value is bool b)
			{
				StringValue = b.ToString().ToLower();
				Cell.CellValue = new CellValue(StringValue);
				Cell.DataType = CellValues.Boolean;
			}
			else if (Value is DateTime dt)
			{
				Cell.CellValue = new CellValue(dt.ToOADate());
				Cell.StyleIndex = 3;    // Date & Time format
				Cell.DataType = CellValues.Number;
				StringValue = dt.ToString("yyyy-MM-dd HH:mm:ss");
			}
			else if (Value is DateTimeOffset dto)
			{
				dt = dto.DateTime;
				Cell.CellValue = new CellValue(dt.ToOADate());
				Cell.StyleIndex = 3;    // Date & Time format
				Cell.DataType = CellValues.Number;
				StringValue = dt.ToString("yyyy-MM-dd HH:mm:ss");
			}
			else if (Value is BlankNode BlankNode)
			{
				StringValue = Result?.GetShortBlankNodeLabel(BlankNode)
					?? BlankNode.ToString();

				Cell.CellValue = new CellValue(StringValue);
				Cell.DataType = CellValues.String;
			}
			else if (Value is UriNode UriNode)
			{
				StringValue = Result?.GetShortUri(UriNode) ?? UriNode.Uri.ToString();

				Cell.CellValue = new CellValue(StringValue);
				Cell.DataType = CellValues.String;
			}
			else
			{
				Cell.CellValue = new CellValue(StringValue);
				Cell.DataType = CellValues.String;
			}

			return Cell;
		}

		private static Stylesheet CreateStylesheet()
		{
			Stylesheet Stylesheet = new Stylesheet();

			NumberingFormats NumberingFormats = new NumberingFormats();
			NumberingFormat DateTimeFormat = new NumberingFormat()
			{
				NumberFormatId = 164U,
				FormatCode = DocumentFormat.OpenXml.StringValue.FromString("yyyy-mm-dd hh:mm:ss")
			};

			NumberingFormats.Append(DateTimeFormat);
			NumberingFormats.Count = 1;
			Stylesheet.Append(NumberingFormats);

			Fonts Fonts = new Fonts();
			Fonts.Append(new Font(                         // Index 0 - default font
				new FontName() { Val = "Calibri" },
				new FontSize() { Val = 11 }));
			Fonts.Append(new Font(                         // Index 1 - bold font
				new FontName() { Val = "Calibri" },
				new FontSize() { Val = 11 },
				new Bold()));
			Fonts.Count = 2;
			Stylesheet.Append(Fonts);

			Fills Fills = new Fills();
			Fills.Append(new Fill // index 0 = none
			{
				PatternFill = new PatternFill
				{
					PatternType = PatternValues.None
				}
			});

			Fills.Append(new Fill // index 1 = Gray125 (required default fills)
			{
				PatternFill = new PatternFill
				{
					PatternType = PatternValues.Gray125
				}
			});
			Fills.Count = 2;
			Stylesheet.Append(Fills);

			Borders Borders = new Borders();
			Borders.Append(new Border(  // index 0 = default (no border)
				new LeftBorder(),
				new RightBorder(),
				new TopBorder(),
				new BottomBorder(),
				new DiagonalBorder()));
			Borders.Count = 1;
			Stylesheet.Append(Borders);

			CellStyleFormats CellStyleFormats = new CellStyleFormats();
			CellStyleFormats.Append(new CellFormat()
			{
				FontId = 0,
				FillId = 0,
				BorderId = 0,
				NumberFormatId = 0
			});
			CellStyleFormats.Count = 1;
			Stylesheet.Append(CellStyleFormats);

			CellFormats CellFormats = new CellFormats();
			CellFormats.Append(new CellFormat()     // index 0 = default
			{
				FontId = 0,
				FillId = 0,
				BorderId = 0
			});
			CellFormats.Append(new CellFormat()     // index 1 = bold
			{
				FontId = 1,
				FillId = 0,
				BorderId = 0,
				ApplyFont = true,
				Alignment = new Alignment()
				{
					Horizontal = HorizontalAlignmentValues.Center
				}
			});
			CellFormats.Append(new CellFormat()     // index 2 = wrap text
			{
				FontId = 0,
				FillId = 0,
				BorderId = 0,
				Alignment = new Alignment()
				{
					WrapText = true,
					Vertical = VerticalAlignmentValues.Top
				}
			});
			CellFormats.Append(new CellFormat()     // index 3 = Date & Time
			{
				FontId = 0,
				FillId = 0,
				BorderId = 0,
				FormatId = 0,
				NumberFormatId = 164U,
				ApplyNumberFormat = true,
			});
			CellFormats.Count = 4;
			Stylesheet.Append(CellFormats);
			
			CellStyles CellStyles = new CellStyles();
			CellStyles.Append(new CellStyle()
			{
				Name = "Normal",
				FormatId = 0,
				BuiltinId = 0
			});
			CellStyles.Count = 1;
			Stylesheet.Append(CellStyles);

			return Stylesheet;
		}

		private static string GetCellReference(uint ColumnIndex, uint RowIndex)
		{
			return GetColumnName(ColumnIndex) + RowIndex.ToString();
		}

		private static string GetColumnName(uint ColumnIndex)
		{
			uint Dividend = ColumnIndex;
			string ColumnName = string.Empty;

			while (Dividend > 0)
			{
				uint Modulo = (Dividend - 1) % 26;
				char Letter = (char)('A' + Modulo);
				ColumnName = Letter + ColumnName;
				Dividend = (Dividend - Modulo - 1) / 26;
			}

			return ColumnName;
		}

		#endregion

		#region Conversion of SPARQL Result sets to Excel spreadsheets

		/// <summary>
		/// Converts a SPARQL Result Set to an Excel spreadsheet.
		/// </summary>
		/// <param name="ResultSet">SPARQL result set.</param>
		/// <param name="FilePath">Path where Excel file will be stored.</param>
		/// <param name="SheetName">Name of sheet.</param>
		/// <returns>Spreadsheet document.</returns>
		public static void ConvertResultSetToExcel(SparqlResultSet ResultSet, string FilePath, string SheetName)
		{
			if (ResultSet is null)
				throw new ArgumentNullException(nameof(ResultSet));

			using SpreadsheetDocument Document = SpreadsheetDocument.Create(FilePath, SpreadsheetDocumentType.Workbook);

			WorkbookPart WorkbookPart = Document.AddWorkbookPart();
			WorkbookPart.Workbook = new Workbook();

			WorksheetPart WorksheetPart = WorkbookPart.AddNewPart<WorksheetPart>();
			WorksheetPart.Worksheet = new Worksheet();

			WorkbookStylesPart StylesPart = WorkbookPart.AddNewPart<WorkbookStylesPart>();
			StylesPart.Stylesheet = CreateStylesheet();  // Define bold and normal styles
			StylesPart.Stylesheet.Save();

			Columns SheetColumns = new Columns();
			WorksheetPart.Worksheet.Append(SheetColumns);
			List<Column> SheetColumn = new List<Column>();

			SheetData SheetData = new SheetData();
			WorksheetPart.Worksheet.Append(SheetData);

			uint[] Increments = ColumnIncrements(ResultSet);
			uint ColumnTotal = ColumnCount(Increments);
			int Columns = ResultSet.Variables?.Length ?? 0;
			int Rows = ResultSet.Records?.Length ?? 0;
			int Row, Column;
			uint x = 1, y = 1;  // Excel uses 1-based indices.
			Row ExcelRow;
			Cell Cell;
			string s;

			if (ResultSet.BooleanResult.HasValue)
			{
				ExcelRow = new Row()
				{
					RowIndex = y
				};

				Cell = new Cell()
				{
					CellReference = GetCellReference(x, y),
					CellValue = new CellValue(ResultSet.BooleanResult.Value.ToString().ToLower()),
					DataType = CellValues.Boolean,
					StyleIndex = 0  // Normal
				};

				ExcelRow.Append(Cell);

				y++;

				SheetData.Append(ExcelRow);
			}

			if (Columns > 0)
			{
				ExcelRow = new Row()
				{
					RowIndex = y
				};

				for (Column = 0; Column < Columns; Column++)
				{
					s = ResultSet.Variables[Column];
					CheckColumnWidth(s, x, SheetColumns, SheetColumn);

					Cell = new Cell()
					{
						CellReference = GetCellReference(x, y),
						CellValue = new CellValue(s),
						DataType = CellValues.String,
						StyleIndex = 1  // Bold
					};

					ExcelRow.Append(Cell);
					x += Increments[Column];
				}

				x = 1;
				y++;

				SheetData.Append(ExcelRow);
			}

			for (Row = 0; Row < Rows; Row++)
			{
				ISparqlResultRecord Record = ResultSet.Records[Row];

				ExcelRow = new Row()
				{
					RowIndex = y
				};

				for (Column = 0; Column < Columns; Column++)
				{
					object Value = Record[ResultSet.Variables[Column]];

					if (!(Value is null))
					{
						s = Value.ToString();
						ExcelRow.Append(CreateCell(ResultSet, x, y, Value, ref s, out string Unit));
						CheckColumnWidth(s, x, SheetColumns, SheetColumn);

						if (!string.IsNullOrEmpty(Unit))
						{
							s = Unit;
							ExcelRow.Append(CreateCell(ResultSet, x + 1, y, Unit, ref s, out Unit));
							CheckColumnWidth(s, x + 1, SheetColumns, SheetColumn);
						}
					}

					x += Increments[Column];
				}

				x = 1;
				y++;

				SheetData.Append(ExcelRow);
			}

			Sheets Sheets = WorkbookPart.Workbook.AppendChild(new Sheets());
			Sheet Sheet = new Sheet()
			{
				Id = WorkbookPart.GetIdOfPart(WorksheetPart),
				SheetId = 1,
				Name = SheetName
			};
			Sheets.Append(Sheet);

			WorksheetPart.Worksheet.Save();
			WorkbookPart.Workbook.Save();
		}

		private static uint[] ColumnIncrements(SparqlResultSet Result)
		{
			string[] Variables = Result.Variables;
			int i, c = Variables?.Length ?? 0;
			uint[] Increments = new uint[c];

			Array.Fill(Increments, (uint)1);

			if (!(Result.Records is null))
			{
				foreach (ISparqlResultRecord Record in Result.Records)
				{
					for (i = 0; i < c; i++)
					{
						if (Increments[i] == 1)
						{
							ISemanticElement Element = Record[Variables[i]];

							if (Element is ISemanticLiteral Literal &&
								Literal.Value is PhysicalQuantity Quantity)
							{
								Increments[i] = 2;
							}
						}
					}
				}
			}

			return Increments;
		}

		#endregion
	}
}
