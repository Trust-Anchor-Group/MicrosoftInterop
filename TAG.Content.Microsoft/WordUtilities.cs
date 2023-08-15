using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Waher.Content.Markdown;

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

				ExportAsMarkdown(MainDocument.Elements(), Markdown, Style);
			}
		}

		private const string MainNaemspace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

		private static bool ExportAsMarkdown(IEnumerable<OpenXmlElement> Elements,
			StringBuilder Markdown, FormattingStyle Style)
		{
			bool HasText = false;

			foreach (OpenXmlElement Element in Elements)
			{
				if (ExportAsMarkdown(Element, Markdown, Style))
					HasText = true;
			}

			return HasText;
		}

		private static bool ExportAsMarkdown(OpenXmlElement Element, StringBuilder Markdown,
			FormattingStyle Style)
		{
			bool HasText = false;

			Console.Out.WriteLine(Element.LocalName);

			if (Element.NamespaceUri == MainNaemspace)
			{
				switch (Element.LocalName)
				{
					case "body":
						if (Element is Body Body)
							HasText = ExportAsMarkdown(Body.Elements(), Markdown, Style);
						break;

					case "p":
						if (Element is Paragraph Paragraph)
						{
							HasText = ExportAsMarkdown(Paragraph.Elements(), Markdown, Style);

							CloseStyles(Style, Markdown, HasText);
							Markdown.AppendLine();
							Markdown.AppendLine();
						}
						break;

					case "pPr":
						if (Element is ParagraphProperties ParagraphProperties)
						{
							// Ignore. Already processed.
						}
						break;

					case "r":
						if (Element is Run Run)
						{
							CheckStyle(Run, Style, Markdown);
							HasText = ExportAsMarkdown(Run.Elements(), Markdown, Style);
						}
						break;

					case "rPr":
						if (Element is ParagraphMarkRunProperties ParagraphMarkRunProperties)
						{
							// Ignore. Already processed.
						}
						else if (Element is RunProperties RunProperties)
						{
							// Ignore. Already processed.
						}
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
						break;

					case "spacing":
						if (Element is SpacingBetweenLines SpacingBetweenLines)
						{
						}
						break;

					case "jc":
						if (Element is Justification Justification)
						{
						}
						else if (Element is TableJustification TableJustification)
						{
						}
						break;

					case "rFonts":
						if (Element is RunFonts RunFonts)
						{
						}
						break;

					case "b":
						if (Element is Bold Bold)
						{
						}
						break;

					case "bCs":
						if (Element is BoldComplexScript BoldComplexScript)
						{
						}
						break;

					case "sz":
						if (Element is FontSize FontSize)
						{
						}
						break;

					case "szCs":
						if (Element is FontSizeComplexScript FontSizeComplexScript)
						{
						}
						break;

					case "tbl":
						if (Element is Table Table)
						{
						}
						break;

					case "tblPr":
						if (Element is TableProperties TableProperties)
						{
						}
						break;

					case "tblStyle":
						if (Element is TableStyle TableStyle)
						{
						}
						break;

					case "tblW":
						if (Element is TableWidth TableWidth)
						{
						}
						break;

					case "tblLayout":
						if (Element is TableLayout TableLayout)
						{
						}
						break;

					case "tblLook":
						if (Element is TableLook TableLook)
						{
						}
						break;

					case "tblGrid":
						if (Element is TableGrid TableGrid)
						{
						}
						break;

					case "gridCol":
						if (Element is GridColumn GridColumn)
						{
						}
						break;

					case "tr":
						if (Element is TableRow TableRow)
						{
						}
						break;

					case "trPr":
						if (Element is TableRowProperties TableRowProperties)
						{
						}
						break;

					case "tc":
						if (Element is TableCell TableCell)
						{
						}
						break;

					case "tcPr":
						if (Element is TableCellProperties TableCellProperties)
						{
						}
						break;

					case "tcW":
						if (Element is TableCellWidth TableCellWidth)
						{
						}
						break;

					case "tcBorders":
						if (Element is TableCellBorders TableCellBorders)
						{
						}
						break;

					case "top":
						if (Element is TopBorder TopBorder)
						{
						}
						break;

					case "bottom":
						if (Element is BottomBorder BottomBorder)
						{
						}
						break;

					case "br":
						if (Element is Break Break)
						{
						}
						break;

					case "pStyle":
						if (Element is ParagraphStyleId ParagraphStyleId)
						{
						}
						break;

					case "numPr":
						if (Element is NumberingProperties NumberingProperties)
						{
						}
						break;

					case "ilvl":
						if (Element is NumberingLevelReference NumberingLevelReference)
						{
						}
						break;

					case "numId":
						if (Element is NumberingId NumberingId)
						{
						}
						break;

					case "ind":
						if (Element is Indentation Indentation)
						{
						}
						break;

					case "lastRenderedPageBreak":
						if (Element is LastRenderedPageBreak LastRenderedPageBreak)
						{
						}
						break;

					case "i":
						if (Element is Italic Italic)
						{
						}
						break;

					case "iCs":
						if (Element is ItalicComplexScript ItalicComplexScript)
						{
						}
						break;

					case "bookmarkStart":
						if (Element is BookmarkStart BookmarkStart)
						{
						}
						break;

					case "bookmarkEnd":
						if (Element is BookmarkEnd BookmarkEnd)
						{
						}
						break;

					case "tblBorders":
						if (Element is TableBorders TableBorders)
						{
						}
						break;

					case "left":
						if (Element is LeftBorder LeftBorder)
						{
						}
						break;

					case "right":
						if (Element is RightBorder RightBorder)
						{
						}
						break;

					case "insideH":
						if (Element is InsideHorizontalBorder InsideHorizontalBorder)
						{
						}
						break;

					case "insideV":
						if (Element is InsideVerticalBorder InsideVerticalBorder)
						{
						}
						break;

					case "sectPr":
						if (Element is SectionProperties SectionProperties)
						{
						}
						break;

					case "footerReference":
						if (Element is FooterReference FooterReference)
						{
						}
						break;

					case "pgSz":
						if (Element is PageSize PageSize)
						{
						}
						break;

					case "pgMar":
						if (Element is PageMargin PageMargin)
						{
						}
						break;

					case "cols":
						if (Element is Columns Columns)
						{
						}
						break;

					case "docGrid":
						if (Element is DocGrid DocGrid)
						{
						}
						break;
				}
			}

			return HasText;
		}

		private class FormattingStyle
		{
			public bool Bold = false;
			public bool Italic = false;
			public bool Underline = false;
			public bool StrikeThrough = false;
			public bool Insert = false;
			public bool Delete = false;
		}

		private static void CheckStyle(Run Item, FormattingStyle Style, StringBuilder Markdown)
		{
			if (!(Item.RunProperties is null))
			{
				RunProperties P = Item.RunProperties;

				bool Bold = !(P.Bold is null);

				if (Style.Bold != Bold)
				{
					Markdown.Append("**");
					Style.Bold = Bold;
				}

				if (!(P.Italic?.Val is null))
				{
					bool Italic = P.Italic.Val.Value;

					if (Style.Italic != Italic)
					{
						Markdown.Append("*");
						Style.Italic = Italic;
					}
				}

				if (!(P.Underline?.Val is null))
				{
					bool Underline;
					bool Insert;

					switch (P.Underline.Val.Value)
					{
						case UnderlineValues.Single:
						case UnderlineValues.Words:
						case UnderlineValues.Dotted:
						case UnderlineValues.Dash:
						case UnderlineValues.DashLong:
						case UnderlineValues.DotDash:
						case UnderlineValues.DotDotDash:
						case UnderlineValues.Wave:
							Underline = true;
							Insert = false;
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
							Underline = false;
							Insert = true;
							break;

						case UnderlineValues.None:
						default:
							Underline = false;
							Insert = false;
							break;
					}

					if (Style.Underline != Underline)
					{
						Markdown.Append("_");
						Style.Underline = Underline;
					}

					if (Style.Insert != Insert)
					{
						Markdown.Append("__");
						Style.Insert = Insert;
					}
				}

				if (!(P.Strike?.Val is null))
				{
					bool StrikeThrough = P.Strike.Val.Value;

					if (Style.StrikeThrough != StrikeThrough)
					{
						Markdown.Append("~");
						Style.StrikeThrough = StrikeThrough;
					}
				}

				if (!(P.DoubleStrike?.Val is null))
				{
					bool Delete = P.DoubleStrike.Val.Value;

					if (Style.Delete != Delete)
					{
						Markdown.Append("~~");
						Style.Delete = Delete;
					}
				}
			}
		}

		private static void CloseStyles(FormattingStyle Style, StringBuilder Markdown, bool HasText)
		{
			if (Style.Bold)
			{
				AppendWhitespaceIfNoText(ref HasText, Markdown);
				Markdown.Append("**");
				Style.Bold = false;
			}

			if (Style.Italic)
			{
				AppendWhitespaceIfNoText(ref HasText, Markdown);
				Markdown.Append('*');
				Style.Italic = false;
			}

			if (Style.Underline)
			{
				AppendWhitespaceIfNoText(ref HasText, Markdown);
				Markdown.Append('_');
				Style.Underline = false;
			}

			if (Style.StrikeThrough)
			{
				AppendWhitespaceIfNoText(ref HasText, Markdown);
				Markdown.Append('~');
				Style.StrikeThrough = false;
			}

			if (Style.Insert)
			{
				AppendWhitespaceIfNoText(ref HasText, Markdown);
				Markdown.Append("__");
				Style.Insert = false;
			}

			if (Style.Delete)
			{
				AppendWhitespaceIfNoText(ref HasText, Markdown);
				Markdown.Append("~~");
				Style.Delete = false;
			}
		}

		private static void AppendWhitespaceIfNoText(ref bool HasText, StringBuilder Markdown)
		{
			if (!HasText)
			{
				Markdown.Append("  ");
				HasText = false;
			}
		}

		/*
				private static void ExportSectionSeparator(Section Section, StringBuilder Markdown)
				{
					int NrColumns = Section.PageSetup.TextColumns.Count;
					int NrChars = 80 / NrColumns;
					if (NrChars < 5)
						NrChars = 5;

					string s = new string('=', NrChars);
					int i;

					for (i = 0; i < NrColumns; i++)
					{
						if (i > 0)
							Markdown.Append(' ');

						Markdown.Append(s);
					}

					Markdown.AppendLine();
					Markdown.AppendLine();
				}

				private static void ExportParagraphs(Document Doc, Paragraphs Paragraphs,
					StringBuilder Markdown)
				{
					foreach (Paragraph Paragraph in Paragraphs)
					{
						WdParagraphAlignment Alignment = Paragraph.Alignment;
						ExportJustificationPrefix(Alignment, Markdown);

						if (Paragraph.get_Style() is Style Style)
						{
							string StyleName = Style.NameLocal.ToUpper();

							switch (StyleName)
							{
								case "TITLE":
								case "BOOK TITLE":
									ExportHeadingParagraph(Doc, Paragraph, 1, Markdown);
									break;

								case "HEADING":
								case "HEADING 1":
								case "HEAD 1":
								case "H1":
									ExportHeadingParagraph(Doc, Paragraph, 2, Markdown);
									break;

								case "HEADING 2":
								case "HEAD 2":
								case "H2":
									ExportHeadingParagraph(Doc, Paragraph, 3, Markdown);
									break;

								case "HEADING 3":
								case "HEAD 3":
								case "H3":
									ExportHeadingParagraph(Doc, Paragraph, 4, Markdown);
									break;

								case "HEADING 4":
								case "HEAD 4":
								case "H4":
									ExportHeadingParagraph(Doc, Paragraph, 5, Markdown);
									break;

								case "HEADING 5":
								case "HEAD 5":
								case "H5":
									ExportHeadingParagraph(Doc, Paragraph, 6, Markdown);
									break;

								case "HEADING 6":
								case "HEAD 6":
								case "H6":
									ExportHeadingParagraph(Doc, Paragraph, 7, Markdown);
									break;

								case "HEADING 7":
								case "HEAD 7":
								case "H7":
									ExportHeadingParagraph(Doc, Paragraph, 8, Markdown);
									break;

								case "HEADING 8":
								case "HEAD 8":
								case "H8":
									ExportHeadingParagraph(Doc, Paragraph, 9, Markdown);
									break;

								case "HEADING 9":
								case "HEAD 9":
								case "H9":
									ExportHeadingParagraph(Doc, Paragraph, 10, Markdown);
									break;

								case "LIST PARAGRAPH":
									ExportListParagraph(Doc, Paragraph, Markdown);
									break;

								case "QUOTE":
								case "INTENSE QUOTE":
									ExportQuoteParagraph(Doc, Paragraph, Markdown);
									break;

								case "NORMAL":
								case "BODY TEXT":
								case "BODY TEXT 2":
								case "BODY TEXT 3":
								case "SUBTITLE":
								default:
									ExportParagraph(Doc, Paragraph, Markdown);
									break;
							}
						}
						else
							ExportParagraph(Doc, Paragraph, Markdown);

						ExportJustificationSuffix(Alignment, Markdown);

						Markdown.AppendLine();
						Markdown.AppendLine();
					}
				}

				private static void ExportJustificationPrefix(WdParagraphAlignment Alignment,
					StringBuilder Markdown)
				{
					switch (Alignment)
					{
						case WdParagraphAlignment.wdAlignParagraphLeft:
						case WdParagraphAlignment.wdAlignParagraphThaiJustify:
						default:
							break;

						case WdParagraphAlignment.wdAlignParagraphCenter:
							Markdown.Append(">>");
							break;

						case WdParagraphAlignment.wdAlignParagraphRight:
							break;

						case WdParagraphAlignment.wdAlignParagraphJustify:
						case WdParagraphAlignment.wdAlignParagraphDistribute:
						case WdParagraphAlignment.wdAlignParagraphJustifyMed:
						case WdParagraphAlignment.wdAlignParagraphJustifyHi:
						case WdParagraphAlignment.wdAlignParagraphJustifyLow:
							Markdown.Append("<<");
							break;
					}
				}

				private static void ExportJustificationSuffix(WdParagraphAlignment Alignment,
					StringBuilder Markdown)
				{
					switch (Alignment)
					{
						case WdParagraphAlignment.wdAlignParagraphLeft:
						case WdParagraphAlignment.wdAlignParagraphThaiJustify:
						default:
							break;

						case WdParagraphAlignment.wdAlignParagraphCenter:
							Markdown.Append("<<");
							break;

						case WdParagraphAlignment.wdAlignParagraphRight:
							Markdown.Append(">>");
							break;

						case WdParagraphAlignment.wdAlignParagraphJustify:
						case WdParagraphAlignment.wdAlignParagraphDistribute:
						case WdParagraphAlignment.wdAlignParagraphJustifyMed:
						case WdParagraphAlignment.wdAlignParagraphJustifyHi:
						case WdParagraphAlignment.wdAlignParagraphJustifyLow:
							Markdown.Append(">>");
							break;
					}
				}

				private static void ExportHeadingParagraph(Document Doc, Paragraph Paragraph,
					int Level, StringBuilder Markdown)
				{
					Markdown.Append(new string('#', Level));
					Markdown.Append(' ');
					ExportParagraph(Doc, Paragraph, Markdown);
				}

				private static void ExportQuoteParagraph(Document Doc, Paragraph Paragraph,
					StringBuilder Markdown)
				{
					Markdown.Append(">\t");
					ExportParagraph(Doc, Paragraph, Markdown);
				}

				private static void ExportListParagraph(Document Doc, Paragraph Paragraph,
					StringBuilder Markdown)
				{
					if (Paragraph.Range.ListFormat.List is null)
						ExportParagraph(Doc, Paragraph, Markdown);
					else
					{
						int Level = Paragraph.Range.ListFormat.ListLevelNumber;
						int ItemNr = Paragraph.Range.ListFormat.ListValue;

						switch (Paragraph.Range.ListFormat.ListType)
						{
							case WdListType.wdListNoNumbering:
							default:
								ExportParagraph(Doc, Paragraph, Markdown);
								break;

							case WdListType.wdListListNumOnly:
							case WdListType.wdListSimpleNumbering:
							case WdListType.wdListOutlineNumbering:
							case WdListType.wdListMixedNumbering:
								ExportIndentation(Level, Markdown);
								Markdown.Append(ItemNr.ToString());
								Markdown.Append(".\t");
								break;

							case WdListType.wdListBullet:
							case WdListType.wdListPictureBullet:
								ExportIndentation(Level, Markdown);
								Markdown.Append(ItemNr.ToString());
								Markdown.Append("*\t");
								break;
						}
					}
				}

				private static void ExportIndentation(int NrTabs, StringBuilder Markdown)
				{
					while (NrTabs-- > 0)
						Markdown.Append('\t');
				}

				private static void ExportParagraph(Document Doc, Paragraph Paragraph,
					StringBuilder Markdown)
				{
					SpanStyle Style = new SpanStyle();
					Range ParagraphRange = Paragraph.Range;

					if (ParagraphRange.FormattedText.SameRange(ParagraphRange))
						ExportSpanRange(ParagraphRange, Style, Markdown);
					else
					{
						foreach (Range Sentance in ParagraphRange.Sentences)
						{
							if (Sentance.FormattedText.SameRange(Sentance))
								ExportSpanRange(Sentance, Style, Markdown);
							else
							{
								foreach (Range Word in Sentance.Words)
								{
									if (Word.FormattedText.SameRange(Word))
										ExportSpanRange(Word, Style, Markdown);
									else
									{
										foreach (Range Character in Sentance.Characters)
											ExportSpanRange(Character, Style, Markdown);
									}
								}
							}
						}
					}

					CloseStyles(Style, Markdown);
				}

				private static bool SameRange(this Range Range1, Range Range2)
				{
					return Range1.Start == Range2.Start && Range1.End == Range2.End;
				}

		*/
	}
}
