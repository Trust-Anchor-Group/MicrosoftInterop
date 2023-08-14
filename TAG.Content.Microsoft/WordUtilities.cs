using Microsoft.Office.Interop.Word;

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
	}
}
