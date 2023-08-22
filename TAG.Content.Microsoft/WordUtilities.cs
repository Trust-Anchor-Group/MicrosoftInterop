using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Waher.Content;
using Waher.Content.Markdown;
using Waher.Content.Xml;
using Waher.Events;
using MarkdownModel = Waher.Content.Markdown.Model;
using Waher.Runtime.Text;

namespace TAG.Content.Microsoft
{
	/// <summary>
	/// Utilities for interoperation with Microsoft Office Word documents.
	/// </summary>
	public static class WordUtilities
	{
		/// <summary>
		/// Converts a Word document to a Markdown document.
		/// </summary>
		/// <param name="WordFileName">File name of Word document.</param>
		/// <param name="MarkdownFileName">File name of Markdown document.</param>
		public static void ConvertWordToMarkdown(string WordFileName, string MarkdownFileName)
		{
			string Markdown = ExtractAsMarkdown(WordFileName);
			File.WriteAllText(MarkdownFileName, Markdown, utf8BomEncoding);
		}

		private static readonly Encoding utf8BomEncoding = new UTF8Encoding(true);

		/// <summary>
		/// Extracts the contents of a Word file to Markdown.
		/// </summary>
		/// <param name="WordFileName">File name of Word document.</param>
		/// <returns>Markdown.</returns>
		public static string ExtractAsMarkdown(string WordFileName)
		{
			return ExtractAsMarkdown(WordFileName, out _);
		}

		/// <summary>
		/// Extracts the contents of a Word file to Markdown.
		/// </summary>
		/// <param name="WordFileName">File name of Word document.</param>
		/// <param name="Language">Language of document.</param>
		/// <returns>Markdown.</returns>
		public static string ExtractAsMarkdown(string WordFileName, out string Language)
		{
			StringBuilder Markdown = new StringBuilder();
			ExtractAsMarkdown(WordFileName, Markdown, out Language);
			return Markdown.ToString();
		}

		/// <summary>
		/// Extracts the contents of a Word file to Markdown.
		/// </summary>
		/// <param name="WordFileName">File name of Word document.</param>
		/// <param name="Markdown">Markdown will be output here.</param>
		public static void ExtractAsMarkdown(string WordFileName, StringBuilder Markdown)
		{
			ExtractAsMarkdown(WordFileName, Markdown, out _);
		}

		/// <summary>
		/// Extracts the contents of a Word file to Markdown.
		/// </summary>
		/// <param name="WordFileName">File name of Word document.</param>
		/// <param name="Markdown">Markdown will be output here.</param>
		/// <param name="Language">Language of document.</param>
		public static void ExtractAsMarkdown(string WordFileName, StringBuilder Markdown,
			out string Language)
		{
			using (WordprocessingDocument Doc = WordprocessingDocument.Open(WordFileName, false))
			{
				Language = Doc.PackageProperties.Language;
				ExtractAsMarkdown(Doc, WordFileName, Markdown);
			}
		}

		/// <summary>
		/// Extracts the contents of a Word file to Markdown.
		/// </summary>
		/// <param name="Doc">Document to convert</param>
		/// <param name="WordFileName">File name of Word document.</param>
		/// <param name="Markdown">Markdown will be output here.</param>
		public static void ExtractAsMarkdown(WordprocessingDocument Doc, string WordFileName, StringBuilder Markdown)
		{
			Document MainDocument = Doc.MainDocumentPart.Document;
			FormattingStyle Style = new FormattingStyle();
			RenderingState State = new RenderingState()
			{
				Doc = Doc,
				FileName = WordFileName,
				FileSize = GetFileSize(WordFileName)
			};
			int StartLen = Markdown.Length;

			ExportAsMarkdown(MainDocument.Elements(), Markdown, Style, State);

			if (!(State.TableFootnotes is null))
			{
				foreach (KeyValuePair<string, string> P in State.TableFootnotes)
				{
					Markdown.AppendLine();
					Markdown.Append("[^");
					Markdown.Append(P.Key);
					Markdown.Append("]:");

					foreach (string Row in GetRows(P.Value))
					{
						Markdown.Append('\t');
						Markdown.AppendLine(Row);
					}
				}
			}

			if (!(State.Footnotes is null))
			{
				foreach (KeyValuePair<long, KeyValuePair<string, bool>> P in State.Footnotes)
				{
					if (!P.Value.Value)
						continue;

					Markdown.AppendLine();
					Markdown.Append("[^fn");
					Markdown.Append(P.Key.ToString());
					Markdown.Append("]:");

					foreach (string Row in GetRows(P.Value.Key))
					{
						Markdown.Append('\t');
						Markdown.AppendLine(Row);
					}
				}
			}

			if (!(State.Endnotes is null))
			{
				foreach (KeyValuePair<long, KeyValuePair<string, bool>> P in State.Endnotes)
				{
					if (!P.Value.Value)
						continue;

					Markdown.AppendLine();
					Markdown.Append("[^en");
					Markdown.Append(P.Key.ToString());
					Markdown.Append("]:");

					foreach (string Row in GetRows(P.Value.Key))
					{
						Markdown.Append('\t');
						Markdown.AppendLine(Row);
					}
				}
			}

			if (!(State.Sections is null) || !(State.MetaData is null))
			{
				string End = Markdown.ToString();
				string Start;

				if (StartLen == 0)
					Start = string.Empty;
				else
				{
					Start = End.Substring(0, StartLen);
					End = End.Substring(StartLen);
				}

				Markdown.Clear();

				if (!(State.MetaData is null))
				{
					foreach (KeyValuePair<string, string> Header in State.MetaData)
					{
						Markdown.Append(Header.Key);
						Markdown.Append(": ");
						Markdown.AppendLine(Header.Value);
					}

					Markdown.AppendLine();
				}

				Markdown.Append(Start);

				if (!(State.Sections is null))
				{
					foreach (string Section in State.Sections)
						Markdown.Append(Section);
				}

				Markdown.Append(End);
			}

			if (!(State.Unrecognized is null))
			{
				StringBuilder Msg = new StringBuilder();

				Msg.AppendLine("Open XML-document with unrecognized elements converted to Markdown.");
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
				Msg.AppendLine(CapLength(XML.PrettyXml(MainDocument.OuterXml), 256 * 1024));
				Msg.AppendLine("```");
#endif
				Log.Warning(Msg.ToString(), WordFileName);
			}
#if DEBUG
			else
			{
				StringBuilder Msg = new StringBuilder();

				Msg.AppendLine("Open XML-document converted to Markdown.");
				Msg.AppendLine();
				Msg.AppendLine("```");
				Msg.AppendLine(CapLength(XML.PrettyXml(MainDocument.OuterXml), 256 * 1024));
				Msg.AppendLine("```");

				Log.Informational(Msg.ToString(), WordFileName);
			}
#endif
		}

		private static long? GetFileSize(string FileName)
		{
			if (string.IsNullOrEmpty(FileName))
				return null;
			else if (!File.Exists(FileName))
				return null;
			else
			{
				try
				{
					using (FileStream fs = File.OpenRead(FileName))
					{
						return fs.Length;
					}
				}
				catch (Exception)
				{
					return null;
				}
			}
		}

		private static string CapLength(string s, int MaxLength)
		{
			if (s.Length > MaxLength)
				return s.Substring(0, MaxLength) + "...";
			else
				return s;
		}

		private static string[] GetRows(string s)
		{
			return s.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
		}

		private const string Word2006Namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
		private const string Word2010Namespace = "http://schemas.microsoft.com/office/word/2010/wordml";
		private const string Drawing2006Namespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
		private const string Drawing2010Namespace = "http://schemas.microsoft.com/office/drawing/2010/main";
		private const string WordDrawing2010Namespace = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
		private const string WordprocessingDrawing2006Namespace = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
		private const string Shape2010Namespace = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
		private const string Picture2006Namespace = "http://schemas.openxmlformats.org/drawingml/2006/picture";
		private const string Compatibility2006Namespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";
		private const string VmlNamespace = "urn:schemas-microsoft-com:vml";

		private static bool ExportAsMarkdown(IEnumerable<OpenXmlElement> Elements,
			StringBuilder Markdown, FormattingStyle Style, RenderingState State)
		{
			bool HasText = false;

			foreach (OpenXmlElement Element in Elements)
			{
				if (ExportAsMarkdown(Element, Markdown, Style, State))
					HasText = true;
			}

			return HasText;
		}

		private static bool ExportAsMarkdown(OpenXmlElement Element, StringBuilder Markdown,
			FormattingStyle Style, RenderingState State)
		{
			bool HasText = false;

			switch (Element.NamespaceUri)
			{
				case Word2006Namespace:
					switch (Element.LocalName)
					{
						case "body":
							if (Element is Body Body)
								HasText = ExportAsMarkdown(Body.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "p":
							if (Element is Paragraph Paragraph)
							{
								StringBuilder ParagraphContent = new StringBuilder();

								Style.NextNumber();

								if (ExportAsMarkdown(Paragraph.Elements(), ParagraphContent, Style, State))
									HasText = true;

								if (Style.CodeBlock)
								{
									StringBuilder Temp = new StringBuilder();

									Temp.AppendLine("```");
									Temp.AppendLine(ParagraphContent.ToString());
									Temp.Append("```");

									ParagraphContent = Temp;

									Style.CodeBlock = false;
									Style.InlineCode = false;
								}

								if (Style.ParagraphType == ParagraphType.Continuation)
								{
									StringBuilder Temp = new StringBuilder();
									bool First = true;

									foreach (string Row in GetRows(ParagraphContent.ToString()))
									{
										if (First)
											First = false;
										else
											Temp.AppendLine();

										Indentation(Temp, Style.PrevItemNumbers?.Count ?? 0);
										Temp.Append(Row);
									}

									ParagraphContent = Temp;
									Style.PrevNumber();
								}

								if (State.Table is null)
								{
									switch (Style.ParagraphAlignment)
									{
										case ParagraphAlignment.Right:
											StringBuilder Temp = new StringBuilder();
											bool First = true;

											foreach (string Row in GetRows(ParagraphContent.ToString()))
											{
												if (First)
													First = false;
												else
													Temp.AppendLine();

												Temp.Append(Row);
												Temp.Append(">>");
											}

											ParagraphContent = Temp;
											break;

										case ParagraphAlignment.Center:
											Temp = new StringBuilder();
											First = true;

											foreach (string Row in GetRows(ParagraphContent.ToString()))
											{
												if (First)
													First = false;
												else
													Temp.AppendLine();

												Markdown.Append(">>");
												Markdown.Append(Row);
												Markdown.Append("<<");
											}

											ParagraphContent = Temp;
											break;

										case ParagraphAlignment.Justified:
											Temp = new StringBuilder();
											First = true;

											foreach (string Row in GetRows(ParagraphContent.ToString()))
											{
												if (First)
													First = false;
												else
													Temp.AppendLine();

												Markdown.Append("<<");
												Markdown.Append(Row);
												Markdown.Append(">>");
											}

											ParagraphContent = Temp;
											break;
									}
								}

								Markdown.AppendLine(ParagraphContent.ToString());

								if (!(State.Table is null))
								{
									if (State.Table.ColumnIndex < State.Table.NrColumns)
									{
										switch (Style.ParagraphAlignment)
										{
											case ParagraphAlignment.Left:
											case ParagraphAlignment.Justified:
											default:
												State.Table.ColumnAlignments[State.Table.ColumnIndex] = MarkdownModel.TextAlignment.Left;
												break;

											case ParagraphAlignment.Right:
												State.Table.ColumnAlignments[State.Table.ColumnIndex] = MarkdownModel.TextAlignment.Right;
												break;

											case ParagraphAlignment.Center:
												State.Table.ColumnAlignments[State.Table.ColumnIndex] = MarkdownModel.TextAlignment.Center;
												break;
										}
									}
								}

								Style.ParagraphAlignment = ParagraphAlignment.Left;
								Markdown.AppendLine();

								if (Style.HorizontalSeparator)
								{
									Markdown.AppendLine(new string('-', 80));
									Markdown.AppendLine();
									Style.HorizontalSeparator = false;
								}

								if (Style.NewSection.HasValue)
								{
									StringBuilder Section = new StringBuilder();

									int NrColumns = Style.NewSection.Value;
									Style.NewSection = null;

									int NrChars = 80 / NrColumns;
									if (NrChars < 5)
										NrChars = 5;

									string s = new string('=', NrChars);
									int i;

									for (i = 0; i < NrColumns; i++)
									{
										if (i > 0)
											Section.Append(' ');

										Section.Append(s);
									}

									Section.AppendLine();
									Section.AppendLine();
									Section.Append(Markdown.ToString());
									Markdown.Clear();

									if (State.Sections is null)
										State.Sections = new LinkedList<string>();

									State.Sections.AddLast(Section.ToString());
									Style.NewSection = null;
								}
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "pPr":
							if (Element is ParagraphProperties ParagraphProperties)
							{
								if (!(ParagraphProperties.Justification is null) &&
									ParagraphProperties.Justification.Val.HasValue)
								{
									switch (ParagraphProperties.Justification.Val.Value)
									{
										case JustificationValues.Left:
										case JustificationValues.Start:
											Style.ParagraphAlignment = ParagraphAlignment.Left;
											break;

										case JustificationValues.Center:
											Style.ParagraphAlignment = ParagraphAlignment.Center;
											break;

										case JustificationValues.Right:
										case JustificationValues.End:
											Style.ParagraphAlignment = ParagraphAlignment.Right;
											break;

										case JustificationValues.Both:
										case JustificationValues.Distribute:
											Style.ParagraphAlignment = ParagraphAlignment.Justified;
											break;

										case JustificationValues.MediumKashida:
										case JustificationValues.NumTab:
										case JustificationValues.HighKashida:
										case JustificationValues.LowKashida:
										case JustificationValues.ThaiDistribute:
										default:
											break;
									}
								}

								Style.ParagraphStyle = true;
								HasText = ExportAsMarkdown(ParagraphProperties.Elements(), Markdown, Style, State);
								Style.ParagraphStyle = false;
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "r":
							if (Element is Run Run)
							{
								HasText = ExportAsMarkdown(Run.Elements(), Markdown, Style, State);

								if (!(Style.StyleChanges is null))
								{
									foreach (char ch in Style.StyleChanges)
									{
										switch (ch)
										{
											case 'b':
												AppendWhitespaceIfNoText(ref HasText, Markdown);
												Markdown.Append("**");
												Style.Bold = false;
												break;

											case 'i':
												AppendWhitespaceIfNoText(ref HasText, Markdown);
												Markdown.Append('*');
												Style.Italic = false;
												break;

											case 'u':
												AppendWhitespaceIfNoText(ref HasText, Markdown);
												Markdown.Append('_');
												Style.Underline = false;
												break;

											case 's':
												AppendWhitespaceIfNoText(ref HasText, Markdown);
												Markdown.Append('~');
												Style.StrikeThrough = false;
												break;

											case 'I':
												AppendWhitespaceIfNoText(ref HasText, Markdown);
												Markdown.Append("__");
												Style.Insert = false;
												break;

											case 'D':
												AppendWhitespaceIfNoText(ref HasText, Markdown);
												Markdown.Append("~~");
												Style.Delete = false;
												break;

											case '^':
												AppendWhitespaceIfNoText(ref HasText, Markdown);
												Markdown.Append(']');
												Style.Superscript = false;
												break;

											case 'v':
												AppendWhitespaceIfNoText(ref HasText, Markdown);
												Markdown.Append(']');
												Style.Subscript = false;
												break;

											case 'c':
												AppendWhitespaceIfNoText(ref HasText, Markdown);
												Markdown.Append('`');
												Style.InlineCode = false;
												break;
										}
									}

									Style.StyleChanges = null;
								}
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "rPr":
							if (Element is ParagraphMarkRunProperties ParagraphMarkRunProperties)
								HasText = ExportAsMarkdown(ParagraphMarkRunProperties.Elements(), Markdown, Style, State);
							else if (Element is RunProperties RunProperties)
								HasText = ExportAsMarkdown(RunProperties.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "rStyle":
							if (!(Element is RunStyle))
								State.UnrecognizedElement(Element);
							break;

						case "b":
							if (Element is Bold)
							{
								if (!Style.Bold && !Style.ParagraphStyle)
								{
									Markdown.Append("**");
									Style.Bold = true;
									Style.StyleChanged('b');
								}
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "bCs":
							if (!(Element is BoldComplexScript))
								State.UnrecognizedElement(Element);
							break;

						case "i":
							if (Element is Italic)
							{
								if (!Style.Italic && !Style.ParagraphStyle)
								{
									Markdown.Append('*');
									Style.Italic = true;
									Style.StyleChanged('i');
								}
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "iCs":
							if (!(Element is ItalicComplexScript))
								State.UnrecognizedElement(Element);
							break;

						case "strike":
							if (Element is Strike)
							{
								if (!Style.StrikeThrough && !Style.ParagraphStyle)
								{
									Markdown.Append('~');
									Style.StrikeThrough = true;
									Style.StyleChanged('s');
								}
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "dstrike":
							if (Element is DoubleStrike)
							{
								if (!Style.Delete && !Style.ParagraphStyle)
								{
									Markdown.Append("~~");
									Style.Delete = true;
									Style.StyleChanged('D');
								}
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "u":
							if (Element is Underline Underline)
							{
								if (Underline.Val.HasValue && !Style.ParagraphStyle)
								{
									switch (Underline.Val.Value)
									{
										case UnderlineValues.Single:
										case UnderlineValues.Words:
										case UnderlineValues.Dotted:
										case UnderlineValues.Dash:
										case UnderlineValues.DashLong:
										case UnderlineValues.DotDash:
										case UnderlineValues.DotDotDash:
										case UnderlineValues.Wave:
										default:
											if (!Style.Underline)
											{
												Markdown.Append('_');
												Style.Underline = true;
												Style.StyleChanged('u');
											}
											break;

										case UnderlineValues.Double:
										case UnderlineValues.Thick:
										case UnderlineValues.DottedHeavy:
										case UnderlineValues.DashedHeavy:
										case UnderlineValues.DashLongHeavy:
										case UnderlineValues.DashDotHeavy:
										case UnderlineValues.DashDotDotHeavy:
										case UnderlineValues.WavyHeavy:
										case UnderlineValues.WavyDouble:
											if (!Style.Insert)
											{
												Markdown.Append("__");
												Style.Insert = true;
												Style.StyleChanged('I');
											}
											break;

										case UnderlineValues.None:
											break;
									}
								}
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "vertAlign":
							if (Element is VerticalTextAlignment VerticalTextAlignment)
							{
								if (VerticalTextAlignment.Val.HasValue)
								{
									switch (VerticalTextAlignment.Val.Value)
									{
										case VerticalPositionValues.Superscript:
											if (!Style.Superscript)
											{
												Markdown.Append("^[");
												Style.Superscript = true;
												Style.StyleChanged('^');
											}
											break;

										case VerticalPositionValues.Subscript:
											if (!Style.Subscript)
											{
												Markdown.Append("[");
												Style.Subscript = true;
												Style.StyleChanged('v');
											}
											break;

										case VerticalPositionValues.Baseline:
											break;
									}
								}
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "sz":
							if (!(Element is FontSize))
								State.UnrecognizedElement(Element);
							break;

						case "szCs":
							if (!(Element is FontSizeComplexScript))
								State.UnrecognizedElement(Element);
							break;

						case "t":
							if (Element is Text Text)
							{
								string s = Text.InnerText;

								if (!string.IsNullOrEmpty(s))
								{
									Markdown.Append(MarkdownDocument.Encode(Text.InnerText));
									HasText = true;
								}
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "tbl":
							if (Element is Table Table)
							{
								TableInfo Bak = State.Table;
								State.Table = new TableInfo();

								HasText = ExportAsMarkdown(Table.Elements(), Markdown, Style, State);
								Markdown.AppendLine();

								State.Table = Bak;
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "tblPr":
							if (Element is TableProperties TableProperties)
								HasText = ExportAsMarkdown(TableProperties.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "tblStyle":
							if (!(Element is TableStyle))
								State.UnrecognizedElement(Element);
							break;

						case "tblW":
							if (!(Element is TableWidth))
								State.UnrecognizedElement(Element);
							break;

						case "tblLayout":
							if (!(Element is TableLayout))
								State.UnrecognizedElement(Element);
							break;

						case "tblLook":
							if (!(Element is TableLook))
								State.UnrecognizedElement(Element);
							break;

						case "tblGrid":
							if (Element is TableGrid TableGrid)
								HasText = ExportAsMarkdown(TableGrid.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "tblHeader":
							if (Element is TableHeader TableHeader)
							{
								if (!(State.Table is null))
								{
									State.Table.IsHeaderRow = true;
									State.Table.HasHeaderRows = true;
								}

								HasText = ExportAsMarkdown(TableHeader.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "gridCol":
							if (Element is GridColumn)
							{
								if (!(State.Table is null))
								{
									State.Table.ColumnAlignments.Add(MarkdownModel.TextAlignment.Left);
									State.Table.NrColumns++;
								}
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "tr":
							if (Element is TableRow TableRow)
							{
								if (!(State.Table is null))
								{
									State.Table.ColumnIndex = 0;
									State.Table.ColumnContents.Clear();
									State.Table.IsHeaderRow = false;
									HasText = ExportAsMarkdown(TableRow.Elements(), Markdown, Style, State);

									int i;

									if (!State.Table.IsHeaderRow && !State.Table.HeaderEmitted)
									{
										State.Table.HeaderEmitted = true;

										if (!State.Table.HasHeaderRows)
										{
											Markdown.Append("| &nbsp; ");

											for (i = 0; i < State.Table.NrColumns; i++)
												Markdown.Append('|');

											Markdown.AppendLine();
										}

										Markdown.Append('|');

										for (i = 0; i < State.Table.NrColumns; i++)
										{
											switch (State.Table.ColumnAlignments[i])
											{
												case MarkdownModel.TextAlignment.Left:
													Markdown.Append(":---|");
													break;

												case MarkdownModel.TextAlignment.Center:
													Markdown.Append(":--:|");
													break;

												case MarkdownModel.TextAlignment.Right:
													Markdown.Append("---:|");
													break;

												default:
													Markdown.Append("----|");
													break;
											}
										}

										Markdown.AppendLine();
									}

									for (i = 0; i < State.Table.NrColumns; i++)
									{
										if (i < State.Table.ColumnContents.Count)
										{
											StringBuilder Column = State.Table.ColumnContents[i];

											if (Column is null)
												Markdown.Append('|');
											else
											{
												Markdown.Append("| ");

												string s = Column.ToString().TrimEnd();
												bool Simple = s.IndexOfAny(simpleCharsProhibited) < 0;

												if (Simple)
												{
													Markdown.Append(s);

													if (!s.EndsWith(" "))
														Markdown.Append(' ');
												}
												else
												{
													if (State.TableFootnotes is null)
														State.TableFootnotes = new Dictionary<string, string>();

													string FootnoteKey = "n" + (State.TableFootnotes.Count + 1).ToString();
													State.TableFootnotes[FootnoteKey] = s;

													Markdown.Append("[^");
													Markdown.Append(FootnoteKey);
													Markdown.Append("] ");
												}
											}
										}
										else
											Markdown.Append('|');
									}

									Markdown.AppendLine("|");
								}
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "trPr":
							if (Element is TableRowProperties TableRowProperties)
								HasText = ExportAsMarkdown(TableRowProperties.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "cnfStyle":
							if (Element is ConditionalFormatStyle ConditionalFormatStyle)
							{
								if (!(State.Table is null) &&
									!State.Table.HasHeaderRows &&
									ConditionalFormatStyle.FirstRow.HasValue &&
									ConditionalFormatStyle.FirstRow.Value)
								{
									State.Table.HasHeaderRows = true;
									State.Table.IsHeaderRow = true;
								}

								HasText = ExportAsMarkdown(ConditionalFormatStyle.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "tc":
							if (Element is TableCell TableCell)
							{
								StringBuilder CellMarkdown = new StringBuilder();
								State.Table.ColumnContents.Add(CellMarkdown);
								HasText = ExportAsMarkdown(TableCell.Elements(), CellMarkdown, Style, State);
								State.Table.ColumnIndex++;
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "tcPr":
							if (Element is TableCellProperties TableCellProperties)
								HasText = ExportAsMarkdown(TableCellProperties.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "tcW":
							if (!(Element is TableCellWidth))
								State.UnrecognizedElement(Element);
							break;

						case "gridSpan":
							if (Element is GridSpan GridSpan)
							{
								HasText = ExportAsMarkdown(GridSpan.Elements(), Markdown, Style, State);

								if (GridSpan.Val.HasValue && !(State.Table is null))
								{
									int i = GridSpan.Val.Value;

									while (i-- > 1)
									{
										State.Table.ColumnContents.Add(null);
										State.Table.ColumnIndex++;
									}
								}
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "tcBorders":
							if (!(Element is TableCellBorders))
								State.UnrecognizedElement(Element);
							break;

						case "spacing":
							if (!(Element is SpacingBetweenLines))
								State.UnrecognizedElement(Element);
							break;

						case "jc":
							if (!(Element is Justification) && !(Element is TableJustification))
								State.UnrecognizedElement(Element);
							break;

						case "rFonts":
							if (Element is RunFonts RunFonts)
							{
								if (IsCodeFont(RunFonts))
								{
									if (Style.ParagraphStyle)
									{
										Style.CodeBlock = true;
										Style.InlineCode = true;
									}
									else if (!Style.InlineCode)
									{
										Markdown.Append('`');
										Style.InlineCode = true;
										Style.StyleChanged('c');
									}
								}
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "top":
							if (!(Element is TopBorder))
								State.UnrecognizedElement(Element);
							break;

						case "bottom":
							if (!(Element is BottomBorder))
								State.UnrecognizedElement(Element);
							break;

						case "br":
							if (Element is Break)
								Markdown.AppendLine("  ");
							else
								State.UnrecognizedElement(Element);
							break;

						case "hyperlink":
							if (Element is Hyperlink Hyperlink)
							{
								Markdown.Append('[');
								HasText = ExportAsMarkdown(Hyperlink.Elements(), Markdown, Style, State);
								Markdown.Append("](");

								if (State.TryGetHyperlink(Hyperlink.Id, out string Link))
									Markdown.Append(Link);

								Markdown.Append(')');
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "pStyle":
							if (Element is ParagraphStyleId ParagraphStyleId)
							{
								string StyleId = ParagraphStyleId.Val?.Value?.ToUpper() ?? string.Empty;

								if (styleIds.TryMap(StyleId, out StyleId))
								{
									switch (StyleId)
									{
										case "TITLE":
											Markdown.Append("# ");
											HasText = true;
											break;

										case "H1":
											Markdown.Append("## ");
											HasText = true;
											break;

										case "H2":
											Markdown.Append("### ");
											HasText = true;
											break;

										case "H3":
											Markdown.Append("#### ");
											HasText = true;
											break;

										case "H4":
											Markdown.Append("##### ");
											HasText = true;
											break;

										case "H5":
											Markdown.Append("###### ");
											HasText = true;
											break;

										case "H6":
											Markdown.Append("####### ");
											HasText = true;
											break;

										case "H7":
											Markdown.Append("######## ");
											HasText = true;
											break;

										case "H8":
											Markdown.Append("######### ");
											HasText = true;
											break;

										case "H9":
											Markdown.Append("########## ");
											HasText = true;
											break;

										case "LIST":
											Style.ParagraphType = ParagraphType.Continuation;
											HasText = true;
											break;

										case "QUOTE":
											Markdown.Append("> ");
											HasText = true;
											break;

										case "NORMAL":
										default:
											HasText = false;
											break;
									}
								}
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "numPr":
							if (Element is NumberingProperties NumberingProperties)
							{
								Style.ItemLevel = null;
								Style.ItemNumber = null;
								HasText = ExportAsMarkdown(NumberingProperties.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "ilvl":
							if (Element is NumberingLevelReference NumberingLevelReference)
							{
								if (NumberingLevelReference.Val.HasValue)
									Style.ItemLevel = NumberingLevelReference.Val.Value;
								else
									Style.ItemLevel = null;

								HasText = ExportAsMarkdown(NumberingLevelReference.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "numId":
							if (Element is NumberingId NumberingId)
							{
								if (!NumberingId.Val.HasValue)
									Style.ItemNumber = null;
								else
								{
									Style.ItemNumber = NumberingId.Val.Value;

									if (Style.ItemLevel.HasValue)
									{
										int i = Style.ItemLevel.Value;

										if (i > 0)
											Indentation(Markdown, i);

										if (State.TryGetNumberingFormat(NumberingId.Val.Value,
											out AbstractNum Num, out _))
										{
											foreach (Level Lvl in Num.Elements<Level>())
											{
												if (i-- == 0)
												{
													if (!(Lvl.StartNumberingValue is null) &&
														Lvl.StartNumberingValue.Val.HasValue)
													{
														Style.OrdinalNumber = Lvl.StartNumberingValue.Val.Value;
													}

													if (!(Lvl.NumberingFormat?.Val is null))
													{

														switch (Lvl.NumberingFormat.Val.Value)
														{
															case NumberFormatValues.Decimal:
															case NumberFormatValues.UpperRoman:
															case NumberFormatValues.LowerRoman:
															case NumberFormatValues.Ordinal:
															case NumberFormatValues.JapaneseCounting:
															case NumberFormatValues.DecimalFullWidth:
															case NumberFormatValues.DecimalHalfWidth:
															case NumberFormatValues.JapaneseLegal:
															case NumberFormatValues.JapaneseDigitalTenThousand:
															case NumberFormatValues.DecimalEnclosedCircle:
															case NumberFormatValues.DecimalFullWidth2:
															case NumberFormatValues.DecimalZero:
															case NumberFormatValues.Ganada:
															case NumberFormatValues.Chosung:
															case NumberFormatValues.DecimalEnclosedFullstop:
															case NumberFormatValues.DecimalEnclosedParen:
															case NumberFormatValues.DecimalEnclosedCircleChinese:
															case NumberFormatValues.TaiwaneseCounting:
															case NumberFormatValues.TaiwaneseCountingThousand:
															case NumberFormatValues.TaiwaneseDigital:
															case NumberFormatValues.ChineseCounting:
															case NumberFormatValues.ChineseCountingThousand:
															case NumberFormatValues.KoreanDigital:
															case NumberFormatValues.KoreanCounting:
															case NumberFormatValues.KoreanLegal:
															case NumberFormatValues.KoreanDigital2:
															case NumberFormatValues.VietnameseCounting:
															case NumberFormatValues.NumberInDash:
															case NumberFormatValues.Hebrew1:
															case NumberFormatValues.ArabicAbjad:
															case NumberFormatValues.HindiNumbers:
															case NumberFormatValues.HindiCounting:
															case NumberFormatValues.ThaiNumbers:
															case NumberFormatValues.ThaiCounting:
																Style.ParagraphType = ParagraphType.OrderedList;
																HasText = true;

																if (Style.SameNubmering || !Style.OrdinalNumber.HasValue)
																	Markdown.Append("#.\t");
																else
																{
																	Markdown.Append(Style.OrdinalNumber.Value.ToString());
																	Markdown.Append(".\t");
																}
																break;

															case NumberFormatValues.UpperLetter:
															case NumberFormatValues.LowerLetter:
															case NumberFormatValues.CardinalText:
															case NumberFormatValues.OrdinalText:
															case NumberFormatValues.Hex:
															case NumberFormatValues.Chicago:
															case NumberFormatValues.IdeographDigital:
															case NumberFormatValues.Aiueo:
															case NumberFormatValues.Iroha:
															case NumberFormatValues.AiueoFullWidth:
															case NumberFormatValues.IrohaFullWidth:
															case NumberFormatValues.Bullet:
															case NumberFormatValues.IdeographEnclosedCircle:
															case NumberFormatValues.IdeographTraditional:
															case NumberFormatValues.IdeographZodiac:
															case NumberFormatValues.IdeographZodiacTraditional:
															case NumberFormatValues.IdeographLegalTraditional:
															case NumberFormatValues.ChineseLegalSimplified:
															case NumberFormatValues.RussianLower:
															case NumberFormatValues.RussianUpper:
															case NumberFormatValues.Hebrew2:
															case NumberFormatValues.ArabicAlpha:
															case NumberFormatValues.HindiVowels:
															case NumberFormatValues.HindiConsonants:
															case NumberFormatValues.ThaiLetters:
															case NumberFormatValues.BahtText:
															case NumberFormatValues.DollarText:
															case NumberFormatValues.Custom:
															default:
																Style.ParagraphType = ParagraphType.BulletList;
																HasText = true;

																if (Style.ItemLevel.HasValue)
																{
																	switch (Style.ItemLevel.Value % 3)
																	{
																		case 0:
																			Markdown.Append("*\t");
																			break;

																		case 1:
																			Markdown.Append("-\t");
																			break;

																		case 2:
																			Markdown.Append("+\t");
																			break;
																	}
																}
																else
																	Markdown.Append("*\t");
																break;

															case NumberFormatValues.None:
																Style.ParagraphType = ParagraphType.Continuation;
																HasText = true;
																Markdown.Append('\t');
																break;
														}
													}
													else if (Lvl.LevelText.Val.HasValue)
													{
														if (Lvl.LevelText.Val.Value.Contains("%"))
															Style.ItemNumber = -1;
														else
															Style.ItemNumber = null;
													}

													break;
												}
											}
										}
									}
								}

								HasText = ExportAsMarkdown(NumberingId.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "ind":
							if (!(Element is Indentation))
								State.UnrecognizedElement(Element);
							break;

						case "lastRenderedPageBreak":
							if (!(Element is LastRenderedPageBreak))
								State.UnrecognizedElement(Element);
							break;

						case "bookmarkStart":
							if (!(Element is BookmarkStart))
								State.UnrecognizedElement(Element);
							break;

						case "bookmarkEnd":
							if (!(Element is BookmarkEnd))
								State.UnrecognizedElement(Element);
							break;

						case "tblBorders":
							if (!(Element is TableBorders))
								State.UnrecognizedElement(Element);
							break;

						case "left":
							if (!(Element is LeftBorder))
								State.UnrecognizedElement(Element);
							break;

						case "right":
							if (!(Element is RightBorder))
								State.UnrecognizedElement(Element);
							break;

						case "insideH":
							if (!(Element is InsideHorizontalBorder))
								State.UnrecognizedElement(Element);
							break;

						case "insideV":
							if (!(Element is InsideVerticalBorder))
								State.UnrecognizedElement(Element);
							break;

						case "sectPr":
							if (Element is SectionProperties SectionProperties)
							{
								HasText = ExportAsMarkdown(SectionProperties.Elements(), Markdown, Style, State);
								if (!Style.NewSection.HasValue)
									Style.NewSection = 1;
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "type":
							if (Element is SectionType SectionType)
							{
								if (SectionType.Val.HasValue)
								{
									switch (SectionType.Val.Value)
									{
										case SectionMarkValues.EvenPage:
										case SectionMarkValues.OddPage:
										case SectionMarkValues.Continuous:
										case SectionMarkValues.NextPage:
											Style.NewSection = 1;
											break;

										case SectionMarkValues.NextColumn:
										default:
											break;
									}
								}

								HasText = ExportAsMarkdown(SectionType.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "col":
							if (!(Element is Column))
								State.UnrecognizedElement(Element);
							break;

						case "pgSz":
							if (Element is PageSize PageSize)
								HasText = ExportAsMarkdown(PageSize.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "pgMar":
							if (Element is PageMargin PageMargin)
								HasText = ExportAsMarkdown(PageMargin.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "cols":
							if (Element is Columns Columns)
							{
								if (Style.NewSection.HasValue && !(Columns.ColumnCount is null) && Columns.ColumnCount.HasValue)
									Style.NewSection = Columns.ColumnCount.Value;

								HasText = ExportAsMarkdown(Columns.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "pBdr":
							if (Element is ParagraphBorders ParagraphBorders)
							{
								if (!(ParagraphBorders.BottomBorder is null))
									Style.HorizontalSeparator = true;
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "footerReference":
							if (!(Element is FooterReference))
								State.UnrecognizedElement(Element);
							break;

						case "docGrid":
							if (!(Element is DocGrid))
								State.UnrecognizedElement(Element);
							break;

						case "lang":
							if (!(Element is Languages))
								State.UnrecognizedElement(Element);
							break;

						case "proofErr":
							if (!(Element is ProofError))
								State.UnrecognizedElement(Element);
							break;

						case "shd":
							if (!(Element is Shading))
								State.UnrecognizedElement(Element);
							break;

						case "color":
							if (!(Element is Color))
								State.UnrecognizedElement(Element);
							break;

						case "kern":
							if (!(Element is Kern))
								State.UnrecognizedElement(Element);
							break;

						case "keepNext":
							if (!(Element is KeepNext))
								State.UnrecognizedElement(Element);
							break;

						case "fldSimple":
							if (Element is SimpleField SimpleField)
							{
								if (!(SimpleField.Instruction is null))
								{
									string Instruction = SimpleField.Instruction.Value;
									Match M = simpleFieldInstruction.Match(Instruction);
									if (M.Success && M.Index == 0 && M.Length == Instruction.Length)
									{
										string Command = M.Groups["Command"].Value;
										string Argument = M.Groups["Argument"].Value;
										//string Type = M.Groups["Type"].Value;
										string Argument2 = M.Groups["Argument2"].Value;
										string Content = null;

										// Ref: http://officeopenxml.com/WPfieldInstructions.php
										switch (Command.ToUpper())
										{
											case "DATE":
											case "TIME":
												Content = ToString(DateTime.Now, Argument2);
												break;

											case "CREATEDATE":
												Content = ToString(State.Doc.PackageProperties.Created, Argument2);
												break;

											case "EDITTIME":
											case "SAVEDATE":
												Content = ToString(State.Doc.PackageProperties.Modified, Argument2);
												break;

											case "PRINTDATE":
												Content = ToString(State.Doc.PackageProperties.LastPrinted, Argument2);
												break;

											case "SUBJECT":
												Content = MarkdownDocument.Encode(State.Doc.PackageProperties.Subject);
												break;

											case "TITLE":
												Content = MarkdownDocument.Encode(State.Doc.PackageProperties.Title);
												break;

											case "REVNUM":
												Content = MarkdownDocument.Encode(State.Doc.PackageProperties.Revision);
												break;

											case "AUTHOR":
												Content = MarkdownDocument.Encode(State.Doc.PackageProperties.Creator);
												break;

											case "LASTSAVEDBY":
												Content = MarkdownDocument.Encode(State.Doc.PackageProperties.LastModifiedBy);
												break;

											case "FILENAME":
												Content = MarkdownDocument.Encode(State.FileName);
												break;

											case "FILESIZE":
												Content = ToString(State.FileSize, Argument2);
												break;

											case "KEYWORDS":
												Content = MarkdownDocument.Encode(State.Doc.PackageProperties.Keywords);
												break;

											case "SEQ":
												if (State.Sequences is null)
													State.Sequences = new Dictionary<string, int>();

												if (!State.Sequences.TryGetValue(Argument, out int i))
													i = 0;

												State.Sequences[Argument] = ++i;
												Content = ToString(i, Argument2);
												break;

											case "FORMCHECKBOX":
											case "FORMDROPDOWN":
											case "FORMTEXT":
											case "TOC":
											case "HYPERLINK":
											case "SECTION":

											case "COMPARE":
											case "DOCVARIABLE":
											case "GOTOBUTTON":
											case "IF":
											case "MACROBUTTON":
											case "PRINT":
											case "COMMENTS":
											case "DOCPROPERTY":
											case "NUMCHARS":
											case "NUMPAGES":
											case "NUMWORDS":
											case "TEMPLATE":
											case "ADVANCE":
											case "SYMBOL":
											case "INDEX":
											case "RD":
											case "TA":
											case "TC":
											case "XE":
											case "AUTOTEXT":
											case "AUTOTEXTLIST":
											case "BIBLIOGRAPHY":
											case "CITATION":
											case "INCLUDEPICTURE":
											case "INCLUDETEXT":
											case "LINK":
											case "NOTEREF":
											case "PAGEREF":
											case "QUOTE":
											case "REF":
											case "STYLEREF":
											case "ADDRESSBLOCK":
											case "ASK":
											case "DATABASE":
											case "FILLIN":
											case "GREETINGLINE":
											case "MERGEFIELD":
											case "MERGEREC":
											case "MERGESEQ":
											case "NEXT":
											case "NEXTIF":
											case "SET":
											case "SKIPIF":
											case "LISTNUM":
											case "PAGE":
											case "SECTIONPAGES":
											case "USERADDRESS":
											case "USERINITIALS":
											case "USERNAME":
												break;
										}

										if (!string.IsNullOrEmpty(Content))
											Markdown.Append(Content);
									}
								}
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "sdt":
							if (Element is SdtBlock SdtBlock)
							{
								Style.DocPartGallery = null;
								HasText = ExportAsMarkdown(SdtBlock.Elements(), Markdown, Style, State);
							}
							else if (Element is SdtRun SdtRun)
							{
								Style.Alias = null;
								Style.ParameterType = ParameterType.String;
								Style.ItemCount = null;

								HasText = ExportAsMarkdown(SdtRun.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "sdtPr":
							if (Element is SdtProperties SdtProperties)
								HasText = ExportAsMarkdown(SdtProperties.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "sdtEndPr":
							if (Element is SdtEndCharProperties SdtEndCharProperties)
								HasText = ExportAsMarkdown(SdtEndCharProperties.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "sdtContent":
							if (Element is SdtContentBlock SdtContentBlock)
							{
								switch (Style.DocPartGallery)
								{
									case "Table of Contents":
										Markdown.AppendLine("![](toc)");
										Markdown.AppendLine();
										HasText = true;
										break;

									default:
										HasText = ExportAsMarkdown(SdtContentBlock.Elements(), Markdown, Style, State);
										break;
								}
							}
							else if (Element is SdtContentRun SdtContentRun)
							{
								if (!string.IsNullOrEmpty(Style.Alias) && Style.ParameterType.HasValue)
								{
									Markdown.Append("[%");
									Markdown.Append(Style.Alias);
									Markdown.Append(']');

									StringBuilder Description = new StringBuilder();
									ExportAsMarkdown(SdtContentRun.Elements(), Description, Style, State);

									State.AddMetaData(Style.Alias, Description.ToString());
									State.AddMetaData(Style.Alias + " Type", Style.ParameterType.Value.ToString());

									HasText = true;
								}
								else
									HasText = ExportAsMarkdown(SdtContentRun.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "id":
							if (!(Element is SdtId))
								State.UnrecognizedElement(Element);
							break;

						case "docPart":
							if (Element is DocPartReference DocPartReference)
								HasText = ExportAsMarkdown(DocPartReference.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "docPartObj":
							if (Element is SdtContentDocPartObject SdtContentDocPartObject)
								HasText = ExportAsMarkdown(SdtContentDocPartObject.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "docPartGallery":
							if (Element is DocPartGallery DocPartGallery)
							{
								Style.DocPartGallery = DocPartGallery.Val?.Value ?? string.Empty;
								HasText = ExportAsMarkdown(DocPartGallery.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "docPartUnique":
							if (!(Element is DocPartUnique))
								State.UnrecognizedElement(Element);
							break;

						case "tabs":
							if (!(Element is Tabs))
								State.UnrecognizedElement(Element);
							break;

						case "noProof":
							if (!(Element is NoProof))
								State.UnrecognizedElement(Element);
							break;

						case "fldChar":
							if (!(Element is FieldChar))
								State.UnrecognizedElement(Element);
							break;

						case "instrText":
							if (!(Element is FieldCode))
								State.UnrecognizedElement(Element);
							break;

						case "webHidden":
							if (!(Element is WebHidden))
								State.UnrecognizedElement(Element);
							break;

						case "tab":
							if (Element is TabChar)
								Markdown.Append('\t');
							else
								State.UnrecognizedElement(Element);
							break;

						case "drawing":
							if (Element is Drawing Drawing)
								HasText = ExportAsMarkdown(Drawing.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "pict":
							if (Element is Picture Picture)
								HasText = ExportAsMarkdown(Picture.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "txbxContent":
							if (Element is TextBoxContent TextBoxContent)
								HasText = ExportAsMarkdown(TextBoxContent.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "footnoteReference":
							if (Element is FootnoteReference FootnoteReference)
							{
								if (FootnoteReference.Id.HasValue &&
									State.TryGetFootnote(FootnoteReference.Id.Value, State, out _))
								{
									Markdown.Append("[^fn");
									Markdown.Append(FootnoteReference.Id.Value.ToString());
									Markdown.Append(']');
									HasText = true;
								}
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "endnoteReference":
							if (Element is EndnoteReference EndnoteReference)
							{
								if (EndnoteReference.Id.HasValue &&
									State.TryGetEndnote(EndnoteReference.Id.Value, State, out _))
								{
									Markdown.Append("[^en");
									Markdown.Append(EndnoteReference.Id.Value.ToString());
									Markdown.Append(']');
									HasText = true;
								}
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "separator":
							if (Element is SeparatorMark)
							{
								Markdown.AppendLine("****************");
								Markdown.AppendLine();
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "continuationSeparator":
							if (Element is ContinuationSeparatorMark)
							{
								Markdown.AppendLine("****************");
								Markdown.AppendLine();
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "footnoteRef":
							if (!(Element is FootnoteReferenceMark))
								State.UnrecognizedElement(Element);
							break;

						case "endnoteRef":
							if (!(Element is EndnoteReferenceMark))
								State.UnrecognizedElement(Element);
							break;

						case "alias":
							if (Element is SdtAlias SdtAlias)
							{
								if (SdtAlias.Val.HasValue)
									Style.Alias = SdtAlias.Val.Value;

								HasText = ExportAsMarkdown(SdtAlias.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "tag":
							if (Element is Tag Tag)
								HasText = ExportAsMarkdown(Tag.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "placeholder":
							if (Element is SdtPlaceholder SdtPlaceholder)
								HasText = ExportAsMarkdown(SdtPlaceholder.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "showingPlcHdr":
							if (Element is ShowingPlaceholder ShowingPlaceholder)
								HasText = ExportAsMarkdown(ShowingPlaceholder.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "dataBinding":
							if (Element is DataBinding DataBinding)
							{
								Style.Alias = null;
								Style.ParameterType = null;
								HasText = ExportAsMarkdown(DataBinding.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "text":
							if (Element is SdtContentText SdtContentText)
								HasText = ExportAsMarkdown(SdtContentText.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "date":
							if (Element is SdtContentDate SdtContentDate)
							{
								Style.ParameterType = ParameterType.DatePicker;
								HasText = ExportAsMarkdown(SdtContentDate.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "lid":
							if (Element is LanguageId LanguageId)
								HasText = ExportAsMarkdown(LanguageId.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "storeMappedDataAs":
							if (Element is SdtDateMappingType SdtDateMappingType)
								HasText = ExportAsMarkdown(SdtDateMappingType.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "calendar":
							if (Element is DocumentFormat.OpenXml.Wordprocessing.Calendar Calendar)
								HasText = ExportAsMarkdown(Calendar.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "comboBox":
							if (Element is SdtContentComboBox SdtContentComboBox)
							{
								Style.ParameterType = ParameterType.ComboBox;
								Style.ItemCount = null;

								HasText = ExportAsMarkdown(SdtContentComboBox.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "dropDownList":
							if (Element is SdtContentDropDownList SdtContentDropDownList)
							{
								Style.ParameterType = ParameterType.ListBox;
								Style.ItemCount = null;

								HasText = ExportAsMarkdown(SdtContentDropDownList.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "dateFormat":
							if (Element is DateFormat DateFormat)
								HasText = ExportAsMarkdown(DateFormat.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "listItem":
							if (Element is ListItem ListItem)
							{
								if (!string.IsNullOrEmpty(ListItem.DisplayText?.Value) &&
									!string.IsNullOrEmpty(ListItem.Value?.Value) &&
									!string.IsNullOrEmpty(Style.Alias) &&
									Style.ParameterType.HasValue)
								{
									int i = Style.ItemCount ?? 0;

									Style.ItemCount = ++i;
									string Key = Style.Alias + " Item" + i.ToString();

									State.AddMetaData(Key + " Value", ListItem.Value.Value);
									State.AddMetaData(Key + " Display", ListItem.DisplayText.Value);
								}

								HasText = ExportAsMarkdown(ListItem.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						default:
							State.UnrecognizedElement(Element);
							break;
					}
					break;

				case Word2010Namespace:
					switch (Element.LocalName)
					{
						case "ligatures":
							if (Element is DocumentFormat.OpenXml.Office2010.Word.Ligatures Ligatures)
								HasText = ExportAsMarkdown(Ligatures.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "checkbox":
							if (Element is DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox SdtContentCheckBox)
							{
								Style.ParameterType = ParameterType.CheckBox;
								HasText = ExportAsMarkdown(SdtContentCheckBox.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "checked":
							if (Element is DocumentFormat.OpenXml.Office2010.Word.Checked Checked)
								HasText = ExportAsMarkdown(Checked.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "checkedState":
							if (Element is DocumentFormat.OpenXml.Office2010.Word.CheckedState CheckedState)
								HasText = ExportAsMarkdown(CheckedState.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "uncheckedState":
							if (Element is DocumentFormat.OpenXml.Office2010.Word.UncheckedState UncheckedState)
								HasText = ExportAsMarkdown(UncheckedState.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						default:
							State.UnrecognizedElement(Element);
							break;
					}
					break;

				case Drawing2006Namespace:
					switch (Element.LocalName)
					{
						case "graphicFrameLocks":
							if (Element is DocumentFormat.OpenXml.Drawing.GraphicFrameLocks GraphicFrameLocks)
								HasText = ExportAsMarkdown(GraphicFrameLocks.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "graphic":
							if (Element is DocumentFormat.OpenXml.Drawing.Graphic Graphic)
								HasText = ExportAsMarkdown(Graphic.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "graphicData":
							if (Element is DocumentFormat.OpenXml.Drawing.GraphicData GraphicData)
								HasText = ExportAsMarkdown(GraphicData.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "blip":
							if (Element is DocumentFormat.OpenXml.Drawing.Blip Blip)
							{
								if (Blip.Embed.HasValue)
								{
									OpenXmlPart Part = State.Doc.MainDocumentPart.GetPartById(Blip.Embed.Value);
									if (Part is ImagePart ImagePart)
									{
										using (Stream ImageStream = ImagePart.GetStream())
										{
											int c = (int)Math.Min(ImageStream.Length, int.MaxValue);
											byte[] Bin = new byte[c];

											ImageStream.Read(Bin, 0, c);

											Markdown.Append("![](data:");
											Markdown.Append(ImagePart.ContentType);
											Markdown.Append(";base64,");
											Markdown.Append(Convert.ToBase64String(Bin));
											Markdown.Append(")");
										}
									}
								}

								HasText = ExportAsMarkdown(Blip.Elements(), Markdown, Style, State);
							}
							else
								State.UnrecognizedElement(Element);
							break;

						case "srcRect":
							if (Element is DocumentFormat.OpenXml.Drawing.SourceRectangle SourceRectangle)
								HasText = ExportAsMarkdown(SourceRectangle.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "stretch":
							if (Element is DocumentFormat.OpenXml.Drawing.Stretch Stretch)
								HasText = ExportAsMarkdown(Stretch.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "xfrm":
							if (Element is DocumentFormat.OpenXml.Drawing.Transform2D Transform2D)
								HasText = ExportAsMarkdown(Transform2D.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "prstGeom":
							if (Element is DocumentFormat.OpenXml.Drawing.PresetGeometry PresetGeometry)
								HasText = ExportAsMarkdown(PresetGeometry.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "noFill":
							if (Element is DocumentFormat.OpenXml.Drawing.NoFill NoFill)
								HasText = ExportAsMarkdown(NoFill.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "ln":
							if (Element is DocumentFormat.OpenXml.Drawing.Outline Outline)
								HasText = ExportAsMarkdown(Outline.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "picLocks":
							if (Element is DocumentFormat.OpenXml.Drawing.PictureLocks PictureLocks)
								HasText = ExportAsMarkdown(PictureLocks.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "extLst":
							if (Element is DocumentFormat.OpenXml.Drawing.BlipExtensionList BlipExtensionList)
								HasText = ExportAsMarkdown(BlipExtensionList.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "fillRect":
							if (Element is DocumentFormat.OpenXml.Drawing.FillRectangle FillRectangle)
								HasText = ExportAsMarkdown(FillRectangle.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "off":
							if (Element is DocumentFormat.OpenXml.Drawing.Offset Offset)
								HasText = ExportAsMarkdown(Offset.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "ext":
							if (Element is DocumentFormat.OpenXml.Drawing.Extents Extents)
								HasText = ExportAsMarkdown(Extents.Elements(), Markdown, Style, State);
							else if (Element is DocumentFormat.OpenXml.Drawing.BlipExtension BlipExtension)
								HasText = ExportAsMarkdown(BlipExtension.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "avLst":
							if (Element is DocumentFormat.OpenXml.Drawing.AdjustValueList AdjustValueList)
								HasText = ExportAsMarkdown(AdjustValueList.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "spLocks":
							if (Element is DocumentFormat.OpenXml.Drawing.ShapeLocks ShapeLocks)
								HasText = ExportAsMarkdown(ShapeLocks.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "solidFill":
							if (Element is DocumentFormat.OpenXml.Drawing.SolidFill SolidFill)
								HasText = ExportAsMarkdown(SolidFill.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "miter":
							if (Element is DocumentFormat.OpenXml.Drawing.Miter Miter)
								HasText = ExportAsMarkdown(Miter.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "headEnd":
							if (Element is DocumentFormat.OpenXml.Drawing.HeadEnd HeadEnd)
								HasText = ExportAsMarkdown(HeadEnd.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "tailEnd":
							if (Element is DocumentFormat.OpenXml.Drawing.TailEnd TailEnd)
								HasText = ExportAsMarkdown(TailEnd.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "spAutoFit":
							if (Element is DocumentFormat.OpenXml.Drawing.ShapeAutoFit ShapeAutoFit)
								HasText = ExportAsMarkdown(ShapeAutoFit.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "srgbClr":
							if (Element is DocumentFormat.OpenXml.Drawing.RgbColorModelHex RgbColorModelHex)
								HasText = ExportAsMarkdown(RgbColorModelHex.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						default:
							State.UnrecognizedElement(Element);
							break;
					}
					break;

				case Drawing2010Namespace:
					switch (Element.LocalName)
					{
						case "useLocalDpi":
							if (Element is DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi UseLocalDpi)
								HasText = ExportAsMarkdown(UseLocalDpi.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						default:
							State.UnrecognizedElement(Element);
							break;
					}
					break;

				case Shape2010Namespace:
					switch (Element.LocalName)
					{
						case "wsp":
							if (Element is DocumentFormat.OpenXml.Office2010.Word.DrawingShape.WordprocessingShape WordprocessingShape)
								HasText = ExportAsMarkdown(WordprocessingShape.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "cNvSpPr":
							if (Element is DocumentFormat.OpenXml.Office2010.Word.DrawingShape.NonVisualDrawingShapeProperties NonVisualDrawingShapeProperties)
								HasText = ExportAsMarkdown(NonVisualDrawingShapeProperties.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "spPr":
							if (Element is DocumentFormat.OpenXml.Office2010.Word.DrawingShape.ShapeProperties ShapeProperties)
								HasText = ExportAsMarkdown(ShapeProperties.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "txbx":
							if (Element is DocumentFormat.OpenXml.Office2010.Word.DrawingShape.TextBoxInfo2 TextBoxInfo2)
								HasText = ExportAsMarkdown(TextBoxInfo2.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "bodyPr":
							if (Element is DocumentFormat.OpenXml.Office2010.Word.DrawingShape.TextBodyProperties TextBodyProperties)
								HasText = ExportAsMarkdown(TextBodyProperties.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						default:
							State.UnrecognizedElement(Element);
							break;
					}
					break;

				case WordDrawing2010Namespace:
					switch (Element.LocalName)
					{
						case "sizeRelH":
							if (Element is DocumentFormat.OpenXml.Office2010.Word.Drawing.RelativeWidth RelativeWidth)
								HasText = ExportAsMarkdown(RelativeWidth.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "sizeRelV":
							if (Element is DocumentFormat.OpenXml.Office2010.Word.Drawing.RelativeHeight RelativeHeight)
								HasText = ExportAsMarkdown(RelativeHeight.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "pctWidth":
							if (Element is DocumentFormat.OpenXml.Office2010.Word.Drawing.PercentageWidth PercentageWidth)
								HasText = ExportAsMarkdown(PercentageWidth.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "pctHeight":
							if (Element is DocumentFormat.OpenXml.Office2010.Word.Drawing.PercentageHeight PercentageHeight)
								HasText = ExportAsMarkdown(PercentageHeight.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						default:
							State.UnrecognizedElement(Element);
							break;
					}
					break;

				case VmlNamespace:
					switch (Element.LocalName)
					{
						case "shapetype":
							if (!(Element is DocumentFormat.OpenXml.Vml.Shapetype))
								State.UnrecognizedElement(Element);
							break;

						case "shape":
							if (!(Element is DocumentFormat.OpenXml.Vml.Shape))
								State.UnrecognizedElement(Element);
							break;

						default:
							State.UnrecognizedElement(Element);
							break;
					}
					break;

				case Picture2006Namespace:
					switch (Element.LocalName)
					{
						case "pic":
							if (Element is DocumentFormat.OpenXml.Drawing.Pictures.Picture Picture)
								HasText = ExportAsMarkdown(Picture.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "nvPicPr":
							if (Element is DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties NonVisualPictureProperties)
								HasText = ExportAsMarkdown(NonVisualPictureProperties.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "blipFill":
							if (Element is DocumentFormat.OpenXml.Drawing.Pictures.BlipFill BlipFill)
								HasText = ExportAsMarkdown(BlipFill.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "spPr":
							if (Element is DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties ShapeProperties)
								HasText = ExportAsMarkdown(ShapeProperties.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "cNvPr":
							if (Element is DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties NonVisualDrawingProperties)
								HasText = ExportAsMarkdown(NonVisualDrawingProperties.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "cNvPicPr":
							if (Element is DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties NonVisualPictureDrawingProperties)
								HasText = ExportAsMarkdown(NonVisualPictureDrawingProperties.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						default:
							State.UnrecognizedElement(Element);
							break;
					}
					break;

				case WordprocessingDrawing2006Namespace:
					switch (Element.LocalName)
					{
						case "inline":
							if (Element is DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline Inline)
								HasText = ExportAsMarkdown(Inline.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "extent":
							if (Element is DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent Extent)
								HasText = ExportAsMarkdown(Extent.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "effectExtent":
							if (Element is DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent EffectExtent)
								HasText = ExportAsMarkdown(EffectExtent.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "docPr":
							if (Element is DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties DocProperties)
								HasText = ExportAsMarkdown(DocProperties.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "cNvGraphicFramePr":
							if (Element is DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties NonVisualGraphicFrameDrawingProperties)
								HasText = ExportAsMarkdown(NonVisualGraphicFrameDrawingProperties.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "anchor":
							if (Element is DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor Anchor)
								HasText = ExportAsMarkdown(Anchor.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "simplePos":
							if (Element is DocumentFormat.OpenXml.Drawing.Wordprocessing.SimplePosition SimplePosition)
								HasText = ExportAsMarkdown(SimplePosition.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "positionH":
							if (Element is DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalPosition HorizontalPosition)
								HasText = ExportAsMarkdown(HorizontalPosition.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "positionV":
							if (Element is DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalPosition VerticalPosition)
								HasText = ExportAsMarkdown(VerticalPosition.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "wrapSquare":
							if (Element is DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapSquare WrapSquare)
								HasText = ExportAsMarkdown(WrapSquare.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "posOffset":
							if (Element is DocumentFormat.OpenXml.Drawing.Wordprocessing.PositionOffset PositionOffset)
								HasText = ExportAsMarkdown(PositionOffset.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "align":
							if (Element is DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignment HorizontalAlignment)
								HasText = ExportAsMarkdown(HorizontalAlignment.Elements(), Markdown, Style, State);
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
								HasText = ExportAsMarkdown(AlternateContent.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "Choice":
							if (Element is AlternateContentChoice AlternateContentChoice)
								HasText = ExportAsMarkdown(AlternateContentChoice.Elements(), Markdown, Style, State);
							else
								State.UnrecognizedElement(Element);
							break;

						case "Fallback":
							if (Element is AlternateContentFallback AlternateContentFallback)
								HasText = ExportAsMarkdown(AlternateContentFallback.Elements(), Markdown, Style, State);
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

		private static string ToString(double? Number, string Argument)
		{
			if (!Number.HasValue)
				return string.Empty;

			return Number.Value.ToString(); // TODO: Output format defined in argument.
		}

		private static string ToString(DateTime? TP, string Argument)
		{
			if (!TP.HasValue)
				return string.Empty;

			try
			{
				return MarkdownDocument.Encode(DateTime.Now.ToString(Argument));
			}
			catch (Exception)
			{
				return MarkdownDocument.Encode(DateTime.Now.ToString());
			}
		}

		private static bool IsCodeFont(RunFonts Fonts)
		{
			return
				monospaceFonts.Contains(Fonts.Ascii?.Value?.ToUpper()) ||
				monospaceFonts.Contains(Fonts.HighAnsi?.Value?.ToUpper()) ||
				monospaceFonts.Contains(Fonts.ComplexScript?.Value?.ToUpper()) ||
				monospaceFonts.Contains(Fonts.EastAsia?.Value?.ToUpper());
		}

		private static void Indentation(StringBuilder Markdown, int NrChars)
		{
			while (NrChars-- > 0)
				Markdown.Append('\t');
		}

		private static readonly Regex simpleFieldInstruction = new Regex(@"^\s*(?'Command'\w+)\s*(?'Argument'[^\\\s]*)\s*(\\(?'Type'[@#*])\s*(?'Argument2'.*))?$", RegexOptions.Singleline | RegexOptions.Compiled);
		private static readonly HarmonizedTextMap styleIds = GetStyleIds();
		private static readonly char[] simpleCharsProhibited = new char[] { '\r', '\n', '|' };
		private static readonly HashSet<string> monospaceFonts = new HashSet<string>()
		{
			"COURIER",
			"COURIER NEW",
			"CONSOLAS",
			"MONACO",
			"LUCIDA CONSOLE",
			"DEJAVU SANS MONO",
			"ROBOTO MONO",
			"SOURCE CODE PRO",
			"FIRA MONO",
			"INCONSOLATA",
			"HACK",
			"MENLO",
			"ANDALE MONO",
			"DROID SANS MONO",
			"TERMIUS",
			"UBUNTU MONO",
			"LIBERATION MONO",
			"ENVY CODE R",
			"FANTASQUE SANS MONO",
			"ANONYMOUS PRO",
			"IBM PLEX MONO",
			"PT MONO"
		};

		private enum ParagraphAlignment
		{
			Left,
			Right,
			Center,
			Justified
		}

		private enum ParagraphType
		{
			Normal,
			BulletList,
			OrderedList,
			Continuation
		}

		private enum ParameterType
		{
			String,
			CheckBox,
			ComboBox,
			ListBox,
			DatePicker
		}

		private class FormattingStyle
		{
			public bool Bold;
			public bool Italic;
			public bool Underline;
			public bool StrikeThrough;
			public bool Insert;
			public bool Delete;
			public bool Superscript;
			public bool Subscript;
			public bool CodeBlock;
			public bool InlineCode;
			public int? NewSection;
			public bool HorizontalSeparator;
			public bool ParagraphStyle;
			public ParagraphType? ParagraphType;
			public int? OrdinalNumber;
			public List<int?> PrevItemNumbers;
			public int? ItemLevel;
			public int? ItemNumber;
			public string DocPartGallery;
			public string Alias;
			public ParameterType? ParameterType;
			public int? ItemCount;
			public ParagraphAlignment ParagraphAlignment;
			public LinkedList<char> StyleChanges;

			public FormattingStyle()
			{
				this.Bold = false;
				this.Italic = false;
				this.Underline = false;
				this.StrikeThrough = false;
				this.Insert = false;
				this.Delete = false;
				this.Superscript = false;
				this.Subscript = false;
				this.InlineCode = false;
				this.CodeBlock = false;
				this.NewSection = null;
				this.ParagraphStyle = false;
				this.ParagraphType = null;
				this.HorizontalSeparator = false;
				this.OrdinalNumber = null;
				this.PrevItemNumbers = null;
				this.ItemLevel = null;
				this.ItemNumber = null;
				this.Alias = null;
				this.ParameterType = null;
				this.ItemCount = null;
				this.ParagraphAlignment = ParagraphAlignment.Left;
				this.StyleChanges = null;
				this.DocPartGallery = null;
			}

			public void StyleChanged(char c)
			{
				if (this.StyleChanges is null)
					this.StyleChanges = new LinkedList<char>();

				this.StyleChanges.AddFirst(c);
			}

			public bool SameNubmering
			{
				get
				{
					if (this.PrevItemNumbers is null ||
						!this.ItemLevel.HasValue ||
						!this.ItemNumber.HasValue)
					{
						return false;
					}

					int Level = this.ItemLevel.Value;
					if (Level >= this.PrevItemNumbers.Count)
						return false;

					int? Number = this.PrevItemNumbers[Level];
					if (!Number.HasValue)
						return false;

					return this.ItemNumber.Value == Number.Value;
				}
			}

			public void NextNumber()
			{
				if (this.ParagraphType == WordUtilities.ParagraphType.Normal)
					this.PrevItemNumbers = null;
				else
				{
					if (this.PrevItemNumbers is null)
						this.PrevItemNumbers = new List<int?>();

					int i = this.ItemLevel ?? 0;
					int c = this.PrevItemNumbers.Count;

					while (c < i - 1)
						this.PrevItemNumbers.Add(null);

					if (c <= i)
						this.PrevItemNumbers.Add(this.ItemNumber);
					else
					{
						this.PrevItemNumbers[i] = this.ItemNumber;
						while (c > i + 1)
							this.PrevItemNumbers.RemoveAt(--c);
					}
				}

				this.ItemLevel = null;
				this.ItemNumber = null;
				this.OrdinalNumber = null;
			}

			public void PrevNumber()
			{
				int i = this.PrevItemNumbers?.Count ?? 0;

				if (i == 0)
				{
					this.ItemNumber = null;
					this.ItemLevel = null;
				}
				else
				{
					this.ItemLevel = --i;
					this.ItemNumber = this.PrevItemNumbers[i];
				}
			}
		}

		private class RenderingState
		{
			public WordprocessingDocument Doc;
			public TableInfo Table = null;
			public Dictionary<string, string> TableFootnotes = null;
			public Dictionary<string, string> Links = null;
			public Dictionary<long, KeyValuePair<string, bool>> Footnotes = null;
			public Dictionary<long, KeyValuePair<string, bool>> Endnotes = null;
			public Dictionary<int, KeyValuePair<AbstractNum, NumberingInstance>> NumberingFormats = null;
			public Dictionary<string, Dictionary<string, int>> Unrecognized = null;
			public Dictionary<string, int> Sequences = null;
			public LinkedList<string> Sections = null;
			public LinkedList<KeyValuePair<string, string>> MetaData = null;
			public string FileName;
			public long? FileSize;

			public void UnrecognizedElement(OpenXmlElement Element)
			{
				if (this.Unrecognized is null)
					this.Unrecognized = new Dictionary<string, Dictionary<string, int>>();

				string Key = Element.LocalName;

				if (Element.NamespaceUri != Word2006Namespace)
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

			public bool TryGetHyperlink(string Id, out string Link)
			{
				if (Id is null)
				{
					Link = null;
					return false;
				}

				if (this.Links is null)
				{
					this.Links = new Dictionary<string, string>();

					foreach (HyperlinkRelationship Rel in this.Doc.MainDocumentPart?.HyperlinkRelationships ?? Array.Empty<HyperlinkRelationship>())
						this.Links[Rel.Id] = Rel.Uri.ToString();
				}

				return this.Links.TryGetValue(Id, out Link);
			}

			public bool TryGetFootnote(long? Id, RenderingState State, out string Content)
			{
				return this.TryGetNote<Footnote>(Id, State, out Content,
					this.Doc.MainDocumentPart?.FootnotesPart.Footnotes,
					ref this.Footnotes);
			}

			public bool TryGetEndnote(long? Id, RenderingState State, out string Content)
			{
				return this.TryGetNote<Endnote>(Id, State, out Content,
					this.Doc.MainDocumentPart?.EndnotesPart.Endnotes,
					ref this.Endnotes);
			}

			private bool TryGetNote<T>(long? Id, RenderingState State, out string Content,
				TypedOpenXmlPartRootElement Root, ref Dictionary<long, KeyValuePair<string, bool>> Notes)
				where T : FootnoteEndnoteType
			{
				if (!Id.HasValue)
				{
					Content = null;
					return false;
				}

				if (Notes is null)
				{
					Notes = new Dictionary<long, KeyValuePair<string, bool>>();

					StringBuilder sb = new StringBuilder();

					foreach (T Note in Root.Elements<T>() ?? Array.Empty<T>())
					{
						if (Note.Id.HasValue)
						{
							FormattingStyle Style = new FormattingStyle();

							sb.Clear();
							ExportAsMarkdown(Note.Elements(), sb, Style, State);

							Notes[Note.Id.Value] = new KeyValuePair<string, bool>(sb.ToString(), false);
						}
					}
				}

				if (!Notes.TryGetValue(Id.Value, out KeyValuePair<string, bool> P))
				{
					Content = null;
					return false;
				}

				Content = P.Key;

				if (!P.Value)
					Notes[Id.Value] = new KeyValuePair<string, bool>(P.Key, true);

				return true;
			}

			public bool TryGetNumberingFormat(int Id, out AbstractNum Numbering,
				out NumberingInstance Instance)
			{
				if (this.NumberingFormats is null)
				{
					Dictionary<int, AbstractNum> Temp = new Dictionary<int, AbstractNum>();

					foreach (AbstractNum N in this.Doc.MainDocumentPart?.NumberingDefinitionsPart.Numbering.Elements<AbstractNum>())
						Temp[N.AbstractNumberId.Value] = N;

					this.NumberingFormats = new Dictionary<int, KeyValuePair<AbstractNum, NumberingInstance>>();

					foreach (NumberingInstance N in this.Doc.MainDocumentPart?.NumberingDefinitionsPart.Numbering.Elements<NumberingInstance>())
					{
						if (N.NumberID.HasValue &&
							!(N.AbstractNumId?.Val is null) &&
							Temp.TryGetValue(N.AbstractNumId.Val.Value, out AbstractNum Num))
						{
							this.NumberingFormats[N.NumberID.Value] =
								new KeyValuePair<AbstractNum, NumberingInstance>(Num, N);
						}
					}
				}

				if (!this.NumberingFormats.TryGetValue(Id, out KeyValuePair<AbstractNum, NumberingInstance> P))
				{
					Numbering = null;
					Instance = null;

					return false;
				}
				else
				{
					Numbering = P.Key;
					Instance = P.Value;
					return true;
				}
			}

			public void AddMetaData(string Key, string Value)
			{
				if (this.MetaData is null)
					this.MetaData = new LinkedList<KeyValuePair<string, string>>();

				foreach (string Row in GetRows(Value.Trim()))
					this.MetaData.AddLast(new KeyValuePair<string, string>(Key, Row));
			}
		}

		private class TableInfo
		{
			public int NrColumns;
			public int ColumnIndex;
			public List<MarkdownModel.TextAlignment> ColumnAlignments = new List<MarkdownModel.TextAlignment>();
			public List<StringBuilder> ColumnContents = new List<StringBuilder>();
			public bool IsHeaderRow;
			public bool HasHeaderRows;
			public bool HeaderEmitted;
		}

		private static HarmonizedTextMap GetStyleIds()
		{
			string Xml = Resources.LoadResourceAsText(typeof(WordUtilities).Namespace + ".StyleMap.xml");
			XmlDocument Doc = new XmlDocument();
			Doc.LoadXml(Xml);

			HarmonizedTextMap Result = new HarmonizedTextMap();

			foreach (XmlNode N in Doc.DocumentElement.ChildNodes)
			{
				if (N is XmlElement E && E.LocalName == "GenericStyleId")
				{
					string To = XML.Attribute(E, "id");

					foreach (XmlNode N2 in E.ChildNodes)
					{
						if (N2 is XmlElement E2 && E2.LocalName == "LocalStyleId")
						{
							string From = XML.Attribute(E2, "id");
							Result.RegisterMapping(From, To);
						}
					}
				}
			}

			return Result;
		}

		private static void AppendWhitespaceIfNoText(ref bool HasText, StringBuilder Markdown)
		{
			if (!HasText)
			{
				Markdown.Append("  ");
				HasText = false;
			}
		}
	}
}
