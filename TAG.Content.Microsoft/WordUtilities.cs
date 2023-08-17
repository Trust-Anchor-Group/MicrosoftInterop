using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using Waher.Content;
using Waher.Content.Markdown;
using Waher.Content.Xml;
using Waher.Events;
using Waher.Networking.HTTP.Vanity;
using MarkdownModel = Waher.Content.Markdown.Model;

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
			StringBuilder Markdown = new StringBuilder();
			ExtractAsMarkdown(WordFileName, Markdown);
			return Markdown.ToString();
		}

		/// <summary>
		/// Extracts the contents of a Word file to Markdown.
		/// </summary>
		/// <param name="WordFileName">File name of Word document.</param>
		/// <param name="Markdown">Markdown will be output here.</param>
		public static void ExtractAsMarkdown(string WordFileName, StringBuilder Markdown)
		{
			using (WordprocessingDocument Doc = WordprocessingDocument.Open(WordFileName, false))
			{
				Document MainDocument = Doc.MainDocumentPart.Document;
				FormattingStyle Style = new FormattingStyle();
				RenderingState State = new RenderingState();
				int StartLen = Markdown.Length;

				ExportAsMarkdown(Doc, MainDocument.Elements(), Markdown, Style, State);

				if (!(State.Footnotes is null))
				{
					foreach (KeyValuePair<string, string> P in State.Footnotes)
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

				if (!(State.Sections is null))
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
					Markdown.Append(Start);
				
					foreach (string Section in State.Sections)
						Markdown.Append(Section);

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

		private const string MainNaemspace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

		private static bool ExportAsMarkdown(WordprocessingDocument Doc, IEnumerable<OpenXmlElement> Elements,
			StringBuilder Markdown, FormattingStyle Style, RenderingState State)
		{
			bool HasText = false;

			foreach (OpenXmlElement Element in Elements)
			{
				if (ExportAsMarkdown(Doc, Element, Markdown, Style, State))
					HasText = true;
			}

			return HasText;
		}

		private static bool ExportAsMarkdown(WordprocessingDocument Doc, OpenXmlElement Element,
			StringBuilder Markdown, FormattingStyle Style, RenderingState State)
		{
			bool HasText = false;

			if (Element.NamespaceUri == MainNaemspace)
			{
				switch (Element.LocalName)
				{
					case "body":
						if (Element is Body Body)
							HasText = ExportAsMarkdown(Doc, Body.Elements(), Markdown, Style, State);
						else
							State.UnrecognizedElement(Element);
						break;

					case "p":
						if (Element is Paragraph Paragraph)
						{
							Style = new FormattingStyle(Style);
							StringBuilder ParagraphContent = new StringBuilder();

							if (!(Paragraph.ParagraphProperties?.ParagraphStyleId is null))
							{
								string StyleId = Paragraph.ParagraphProperties.ParagraphStyleId.Val?.Value?.ToUpper() ?? string.Empty;
								if (styleIds.CheckVanityResource(ref StyleId))
								{
									switch (StyleId)
									{
										case "TITLE":
											ParagraphContent.Append("# ");
											HasText = true;
											break;

										case "H1":
											ParagraphContent.Append("## ");
											HasText = true;
											break;

										case "H2":
											ParagraphContent.Append("### ");
											HasText = true;
											break;

										case "H3":
											ParagraphContent.Append("#### ");
											HasText = true;
											break;

										case "H4":
											ParagraphContent.Append("##### ");
											HasText = true;
											break;

										case "H5":
											ParagraphContent.Append("###### ");
											HasText = true;
											break;

										case "H6":
											ParagraphContent.Append("####### ");
											HasText = true;
											break;

										case "H7":
											ParagraphContent.Append("######## ");
											HasText = true;
											break;

										case "H8":
											ParagraphContent.Append("######### ");
											HasText = true;
											break;

										case "H9":
											ParagraphContent.Append("########## ");
											HasText = true;
											break;

										case "UL":
											ParagraphContent.Append("* ");
											HasText = true;
											break;

										case "OL":
											ParagraphContent.Append("#. ");
											HasText = true;
											break;

										case "QUOTE":
											ParagraphContent.Append("> ");
											HasText = true;
											break;

										case "NORMAL":
										default:
											HasText = false;
											break;
									}
								}
							}

							if (ExportAsMarkdown(Doc, Paragraph.Elements(), ParagraphContent, Style, State))
								HasText = true;

							switch (Style.ParagraphAlignment)
							{
								case ParagraphAlignment.Left:
								default:
									Markdown.AppendLine(ParagraphContent.ToString());
									break;

								case ParagraphAlignment.Right:
									foreach (string Row in GetRows(ParagraphContent.ToString()))
									{
										Markdown.Append(Row);
										Markdown.AppendLine(">>");
									}
									break;

								case ParagraphAlignment.Center:
									foreach (string Row in GetRows(ParagraphContent.ToString()))
									{
										Markdown.Append(">>");
										Markdown.Append(Row);
										Markdown.AppendLine("<<");
									}
									break;

								case ParagraphAlignment.Justified:
									foreach (string Row in GetRows(ParagraphContent.ToString()))
									{
										Markdown.Append("<<");
										Markdown.Append(Row);
										Markdown.AppendLine(">>");
									}
									break;
							}

							Markdown.AppendLine();

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

							HasText = ExportAsMarkdown(Doc, ParagraphProperties.Elements(), Markdown, Style, State);
						}
						else
							State.UnrecognizedElement(Element);
						break;

					case "r":
						if (Element is Run Run)
						{
							FormattingStyle RunStyle = new FormattingStyle(Style);
							HasText = ExportAsMarkdown(Doc, Run.Elements(), Markdown, RunStyle, State);

							if (Style.Bold ^ RunStyle.Bold)
							{
								AppendWhitespaceIfNoText(ref HasText, Markdown);
								Markdown.Append("**");
							}

							if (Style.Italic ^ RunStyle.Italic)
							{
								AppendWhitespaceIfNoText(ref HasText, Markdown);
								Markdown.Append('*');
							}

							if (Style.Underline ^ RunStyle.Underline)
							{
								AppendWhitespaceIfNoText(ref HasText, Markdown);
								Markdown.Append('_');
							}

							if (Style.StrikeThrough ^ RunStyle.StrikeThrough)
							{
								AppendWhitespaceIfNoText(ref HasText, Markdown);
								Markdown.Append('~');
							}

							if (Style.Insert ^ RunStyle.Insert)
							{
								AppendWhitespaceIfNoText(ref HasText, Markdown);
								Markdown.Append("__");
							}

							if (Style.Delete ^ RunStyle.Delete)
							{
								AppendWhitespaceIfNoText(ref HasText, Markdown);
								Markdown.Append("~~");
							}

							if (Style.Superscript ^ RunStyle.Superscript)
							{
								AppendWhitespaceIfNoText(ref HasText, Markdown);
								Markdown.Append(']');
							}

							if (Style.Subscript ^ RunStyle.Subscript)
							{
								AppendWhitespaceIfNoText(ref HasText, Markdown);
								Markdown.Append(']');
							}

							if (Style.Code ^ RunStyle.Code)
							{
								AppendWhitespaceIfNoText(ref HasText, Markdown);
								Markdown.Append('`');
							}
						}
						else
							State.UnrecognizedElement(Element);
						break;

					case "rPr":
						if (Element is ParagraphMarkRunProperties ParagraphMarkRunProperties)
							HasText = ExportAsMarkdown(Doc, ParagraphMarkRunProperties.Elements(), Markdown, Style, State);
						else if (Element is RunProperties RunProperties)
							HasText = ExportAsMarkdown(Doc, RunProperties.Elements(), Markdown, Style, State);
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
							if (!Style.Bold)
							{
								Markdown.Append("**");
								Style.Bold = true;
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
							if (!Style.Italic)
							{
								Markdown.Append('*');
								Style.Italic = true;
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
							if (!Style.StrikeThrough)
							{
								Markdown.Append('~');
								Style.StrikeThrough = true;
							}
						}
						else
							State.UnrecognizedElement(Element);
						break;

					case "dstrike":
						if (Element is DoubleStrike)
						{
							if (!Style.Delete)
							{
								Markdown.Append("~~");
								Style.Delete = true;
							}
						}
						else
							State.UnrecognizedElement(Element);
						break;

					case "u":
						if (Element is Underline Underline)
						{
							if (Underline.Val.HasValue)
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
										}
										break;

									case UnderlineValues.None:
										if (Style.Underline)
										{
											Markdown.Append('_');
											Style.Underline = false;
										}

										if (Style.Insert)
										{
											Markdown.Append("__");
											Style.Insert = false;
										}
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
										}
										break;

									case VerticalPositionValues.Subscript:
										if (!Style.Subscript)
										{
											Markdown.Append("[");
											Style.Subscript = true;
										}
										break;

									case VerticalPositionValues.Baseline:
										if (Style.Superscript)
										{
											Markdown.Append(']');
											Style.Superscript = false;
										}

										if (Style.Subscript)
										{
											Markdown.Append(']');
											Style.Subscript = false;
										}
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

							HasText = ExportAsMarkdown(Doc, Table.Elements(), Markdown, Style, State);
							Markdown.AppendLine();

							State.Table = Bak;
						}
						else
							State.UnrecognizedElement(Element);
						break;

					case "tblPr":
						if (Element is TableProperties TableProperties)
							HasText = ExportAsMarkdown(Doc, TableProperties.Elements(), Markdown, Style, State);
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
							HasText = ExportAsMarkdown(Doc, TableGrid.Elements(), Markdown, Style, State);
						else
							State.UnrecognizedElement(Element);
						break;

					case "gridCol":
						if (Element is GridColumn)
						{
							if (!(State.Table is null))
							{
								State.Table.ColumnAlignments.Add(MarkdownModel.TextAlignment.Left); // TODO
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
								State.Table.ColumnContents.Clear();
								HasText = ExportAsMarkdown(Doc, TableRow.Elements(), Markdown, Style, State);

								int i;

								for (i = 0; i < State.Table.NrColumns; i++)
								{
									Markdown.Append("| ");

									if (i < State.Table.ColumnContents.Count)
									{
										string s = State.Table.ColumnContents[i].ToString();
										bool Simple = s.IndexOfAny(simpleCharsProhibited) < 0;

										if (Simple)
										{
											Markdown.Append(s);

											if (!s.EndsWith(" "))
												Markdown.Append(' ');
										}
										else
										{
											if (State.Footnotes is null)
												State.Footnotes = new Dictionary<string, string>();

											string FootnoteKey = "n" + (++State.NrFootnotes).ToString();
											State.Footnotes[FootnoteKey] = s;

											Markdown.Append("[^");
											Markdown.Append(FootnoteKey);
											Markdown.Append("] ");
										}
									}
								}

								Markdown.AppendLine("|");
							}
						}
						else
							State.UnrecognizedElement(Element);
						break;

					case "trPr":
						if (Element is TableRowProperties TableRowProperties)
							HasText = ExportAsMarkdown(Doc, TableRowProperties.Elements(), Markdown, Style, State);
						else
							State.UnrecognizedElement(Element);
						break;

					case "tc":
						if (Element is TableCell TableCell)
						{
							StringBuilder CellMarkdown = new StringBuilder();
							State.Table.ColumnContents.Add(CellMarkdown);
							HasText = ExportAsMarkdown(Doc, TableCell.Elements(), CellMarkdown, Style, State);
						}
						else
							State.UnrecognizedElement(Element);
						break;

					case "tcPr":
						if (Element is TableCellProperties TableCellProperties)
							HasText = ExportAsMarkdown(Doc, TableCellProperties.Elements(), Markdown, Style, State);
						else
							State.UnrecognizedElement(Element);
						break;

					case "tcW":
						if (!(Element is TableCellWidth))
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
							if (IsCodeFont(RunFonts) ^ Style.Code)
							{
								Markdown.Append('`');
								Style.Code = !Style.Code;
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
							HasText = ExportAsMarkdown(Doc, Hyperlink.Elements(), Markdown, Style, State);
							Markdown.Append("](");

							foreach (HyperlinkRelationship Rel in Doc.MainDocumentPart.HyperlinkRelationships)
							{
								if (Rel.Id == Hyperlink.Id)
								{
									Markdown.Append(Rel.Uri.ToString());
									break;
								}
							}

							Markdown.Append(')');
						}
						else
							State.UnrecognizedElement(Element);
						break;

					case "pStyle":
						if (!(Element is ParagraphStyleId))
							State.UnrecognizedElement(Element);
						break;

					case "numPr":
						if (!(Element is NumberingProperties))
							State.UnrecognizedElement(Element);
						break;

					case "ilvl":
						if (!(Element is NumberingLevelReference))
							State.UnrecognizedElement(Element);
						break;

					case "numId":
						if (!(Element is NumberingId))
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
							HasText = ExportAsMarkdown(Doc, SectionProperties.Elements(), Markdown, Style, State);
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

							HasText = ExportAsMarkdown(Doc, SectionType.Elements(), Markdown, Style, State);
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
							HasText = ExportAsMarkdown(Doc, PageSize.Elements(), Markdown, Style, State);
						else
							State.UnrecognizedElement(Element);
						break;

					case "pgMar":
						if (Element is PageMargin PageMargin)
							HasText = ExportAsMarkdown(Doc, PageMargin.Elements(), Markdown, Style, State);
						else
							State.UnrecognizedElement(Element);
						break;

					case "cols":
						if (Element is Columns Columns)
						{
							if (Style.NewSection.HasValue && !(Columns.ColumnCount is null) && Columns.ColumnCount.HasValue)
								Style.NewSection = Columns.ColumnCount.Value;

							HasText = ExportAsMarkdown(Doc, Columns.Elements(), Markdown, Style, State);
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

					default:
						State.UnrecognizedElement(Element);
						break;
				}
			}

			return HasText;
		}

		private static bool IsCodeFont(RunFonts Fonts)
		{
			return
				monospaceFonts.Contains(Fonts.Ascii?.Value?.ToUpper()) ||
				monospaceFonts.Contains(Fonts.HighAnsi?.Value?.ToUpper()) ||
				monospaceFonts.Contains(Fonts.ComplexScript?.Value?.ToUpper()) ||
				monospaceFonts.Contains(Fonts.EastAsia?.Value?.ToUpper());
		}

		private static readonly VanityResources styleIds = GetStyleIds();
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
			public bool Code;
			public int? NewSection;
			public ParagraphAlignment ParagraphAlignment;

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
				this.Code = false;
				this.NewSection = null;
				this.ParagraphAlignment = ParagraphAlignment.Left;
			}

			public FormattingStyle(FormattingStyle Prev)
			{
				this.Bold = Prev.Bold;
				this.Italic = Prev.Italic;
				this.Underline = Prev.Underline;
				this.StrikeThrough = Prev.StrikeThrough;
				this.Insert = Prev.Insert;
				this.Delete = Prev.Delete;
				this.Superscript = Prev.Superscript;
				this.Subscript = Prev.Subscript;
				this.Code = Prev.Code;
				this.NewSection = Prev.NewSection;
				this.ParagraphAlignment = Prev.ParagraphAlignment;
			}
		}

		private class RenderingState
		{
			public TableInfo Table = null;
			public Dictionary<string, string> Footnotes = null;
			public int NrFootnotes = 0;
			public Dictionary<string, Dictionary<string, int>> Unrecognized = null;
			public LinkedList<string> Sections = null;

			public void UnrecognizedElement(OpenXmlElement Element)
			{
				if (this.Unrecognized is null)
					this.Unrecognized = new Dictionary<string, Dictionary<string, int>>();

				if (!this.Unrecognized.TryGetValue(Element.LocalName, out Dictionary<string, int> ByType))
				{
					ByType = new Dictionary<string, int>();
					this.Unrecognized[Element.LocalName] = ByType;
				}

				string s = Element.GetType().FullName;

				if (!ByType.TryGetValue(s, out int i))
					i = 1;
				else
					i++;

				ByType[s] = i;
			}
		}

		private class TableInfo
		{
			public int NrColumns;
			public List<MarkdownModel.TextAlignment> ColumnAlignments = new List<MarkdownModel.TextAlignment>();
			public List<StringBuilder> ColumnContents = new List<StringBuilder>();
			public bool HeaderRow = true;
		}

		private static VanityResources GetStyleIds()
		{
			string Xml = Resources.LoadResourceAsText(typeof(WordUtilities).Namespace + ".StyleMap.xml");
			XmlDocument Doc = new XmlDocument();
			Doc.LoadXml(Xml);

			VanityResources Result = new VanityResources();

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
							Result.RegisterVanityResource(From, To);
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
