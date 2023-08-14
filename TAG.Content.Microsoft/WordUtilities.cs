using Microsoft.Office.Interop.Word;
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
		/// Converts a Word document to a PDF document.
		/// </summary>
		/// <param name="WordFileName">File name of Word document.</param>
		/// <param name="PdfFileName">File name of PDF document.</param>
		public static void ConvertWordToPdf(string WordFileName, string PdfFileName)
		{
			ConvertWordToPdf(WordFileName, PdfFileName, false);
		}

		/// <summary>
		/// Converts a Word document to a PDF document.
		/// </summary>
		/// <param name="WordFileName">File name of Word document.</param>
		/// <param name="PdfFileName">File name of PDF document.</param>
		/// <param name="ForPrint">If the conversion is for print (true) or screen (false, and default).</param>
		public static void ConvertWordToPdf(string WordFileName, string PdfFileName, bool ForPrint)
		{
			Application Word = new Application();
			try
			{
				Document Doc = Word.Documents.Open(
					FileName: WordFileName,
					ConfirmConversions: false,
					ReadOnly: true,
					AddToRecentFiles: false);
				try
				{
					Doc.ExportAsFixedFormat(
						OutputFileName: PdfFileName,
						ExportFormat: WdExportFormat.wdExportFormatPDF,
						OpenAfterExport: false,
						OptimizeFor: ForPrint ? WdExportOptimizeFor.wdExportOptimizeForPrint : WdExportOptimizeFor.wdExportOptimizeForOnScreen,
						Range: WdExportRange.wdExportAllDocument,
						Item: WdExportItem.wdExportDocumentContent);
				}
				finally
				{
					Doc.Close(SaveChanges: false);
				}
			}
			finally
			{
				Word.Quit();
			}
		}

		/// <summary>
		/// Converts a Word document to a XPS document.
		/// </summary>
		/// <param name="WordFileName">File name of Word document.</param>
		/// <param name="XpsFileName">File name of XPS document.</param>
		public static void ConvertWordToXps(string WordFileName, string XpsFileName)
		{
			Application Word = new Application();
			try
			{
				Document Doc = Word.Documents.Open(
					FileName: WordFileName,
					ConfirmConversions: false,
					ReadOnly: true,
					AddToRecentFiles: false);
				try
				{
					Doc.ExportAsFixedFormat(
						OutputFileName: XpsFileName,
						ExportFormat: WdExportFormat.wdExportFormatXPS,
						OpenAfterExport: false,
						Range: WdExportRange.wdExportAllDocument,
						Item: WdExportItem.wdExportDocumentContent);
				}
				finally
				{
					Doc.Close(SaveChanges: false);
				}
			}
			finally
			{
				Word.Quit();
			}
		}

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
			Application Word = new Application();
			try
			{
				Document Doc = Word.Documents.Open(
					FileName: WordFileName,
					ConfirmConversions: false,
					ReadOnly: true,
					AddToRecentFiles: false);
				try
				{
					/***********************
					 * Doc.Bibliography;
					 * Doc.Characters;
					 * Doc.Content;
					 * Doc.Fields;
					 * Doc.Footnotes;
					 * Doc.FormFields;
					 * Doc.Frames;
					 * Doc.Frameset;
					 * Doc.Hyperlinks;
					 * Doc.Indexes;
					 * Doc.InlineShapes;
					 * Doc.ListParagraphs;
					 * Doc.Lists;
					 * Doc.Paragraphs;
					 * Doc.Shapes;
					 * Doc.Tables;
					 * Doc.TablesOfContents;
					 * Doc.TablesOfFigures;
					 * Doc.Variables;
					 *************************/

					foreach (Range Item in Doc.StoryRanges)
					{
						switch (Item.StoryType)
						{
							case WdStoryType.wdMainTextStory:
								ExportSectionSeparator(Item.Sections.First, Markdown);
								ExportParagraphs(Doc, Item.Paragraphs, Markdown);
								break;

							case WdStoryType.wdFootnotesStory:
							case WdStoryType.wdEndnotesStory:
							case WdStoryType.wdCommentsStory:
							case WdStoryType.wdTextFrameStory:
							case WdStoryType.wdEvenPagesHeaderStory:
							case WdStoryType.wdPrimaryHeaderStory:
							case WdStoryType.wdEvenPagesFooterStory:
							case WdStoryType.wdPrimaryFooterStory:
							case WdStoryType.wdFirstPageHeaderStory:
							case WdStoryType.wdFirstPageFooterStory:
							case WdStoryType.wdFootnoteSeparatorStory:
							case WdStoryType.wdFootnoteContinuationSeparatorStory:
							case WdStoryType.wdFootnoteContinuationNoticeStory:
							case WdStoryType.wdEndnoteSeparatorStory:
							case WdStoryType.wdEndnoteContinuationSeparatorStory:
							case WdStoryType.wdEndnoteContinuationNoticeStory:
								break;
						}
					}
				}
				finally
				{
					Doc.Close(SaveChanges: false);
				}
			}
			finally
			{
				Word.Quit();
			}
		}

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

		private static void ExportSpanRange(Range Span, SpanStyle Style, StringBuilder Markdown)
		{
			CheckStyle(Span, Style, Markdown);
			Markdown.Append(MarkdownDocument.Encode(Span.Text));
		}

		private static void CheckStyle(Range Item, SpanStyle Style, StringBuilder Markdown)
		{
			Font Font = Item.Font;

			if ((Style.Bold && Font.Bold == 0) || (!Style.Bold && Font.Bold != 0))
			{
				Markdown.Append("**");
				Style.Bold = !Style.Bold;
			}

			if ((Style.Italic && Font.Italic == 0) || (!Style.Italic && Font.Italic != 0))
			{
				Markdown.Append('*');
				Style.Italic = !Style.Italic;
			}

			if ((Style.Underline && Font.Underline == WdUnderline.wdUnderlineNone) ||
				(!Style.Underline && Font.Underline != WdUnderline.wdUnderlineNone))
			{
				Markdown.Append('_');
				Style.Underline = !Style.Underline;
			}

			if ((Style.StrikeThrough && Font.StrikeThrough == 0) ||
				(!Style.StrikeThrough && Font.StrikeThrough != 0))
			{
				Markdown.Append('~');
				Style.StrikeThrough = !Style.StrikeThrough;
			}
		}

		private static void CloseStyles(SpanStyle Style, StringBuilder Markdown)
		{
			if (Style.Bold)
			{
				Markdown.Append("**");
				Style.Bold = false;
			}

			if (Style.Italic)
			{
				Markdown.Append('*');
				Style.Italic = false;
			}

			if (Style.Underline)
			{
				Markdown.Append('_');
				Style.Underline = false;
			}

			if (Style.StrikeThrough)
			{
				Markdown.Append('~');
				Style.StrikeThrough = false;
			}
		}

		private class SpanStyle
		{
			public bool Bold = false;
			public bool Italic = false;
			public bool Underline = false;
			public bool StrikeThrough = false;
		}
	}
}
