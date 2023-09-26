using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using TAG.Content.Microsoft;
using Waher.Content;
using Waher.Content.Markdown;
using Waher.Runtime.Inventory;

namespace WordToMarkdown
{
	/// <summary>
	/// Interaction logic for App.xaml
	/// </summary>
	public partial class App : Application
	{
		protected override void OnStartup(StartupEventArgs e)
		{
			Types.Initialize(
				typeof(App).Assembly,
				typeof(InternetContent).Assembly,
				typeof(WordUtilities).Assembly);

			base.OnStartup(e);

			int i = 0;
			int c = e.Args.Length;
			bool Recursive = false;

			if (c == 0)
			{
				MainWindow MainWindow = new();
				MainWindow.Show();
			}
			else
			{
				List<KeyValuePair<string, string>>? Headers = null;
				string? InputFileName = null;
				string? OutputFileName = null;
				bool Error = false;

				while (i < c)
				{
					switch (e.Args[i++].ToLower())
					{
						case "-i":
						case "-input":
						case "-word":
							if (InputFileName is null)
							{
								if (i < c)
									InputFileName = e.Args[i++];
								else
								{
									Error = true;
									Console.Error.WriteLine("Missing input file name.");
								}
							}
							else
							{
								Error = true;
								Console.Error.WriteLine("Input file name already provided.");
							}
							break;

						case "-o":
						case "-output":
						case "-md":
						case "-markdown":
							if (OutputFileName is null)
							{
								if (i < c)
									OutputFileName = e.Args[i++];
								else
								{
									Error = true;
									Console.Error.WriteLine("Missing output file name.");
								}
							}
							else
							{
								Error = true;
								Console.Error.WriteLine("Output file name already provided.");
							}
							break;

						case "-meta":
						case "-header":
							if (i < c)
							{
								string s = e.Args[i++];
								int j = s.IndexOf('=');

								if (j < 0)
								{
									Error = true;
									Console.Out.WriteLine("Invalid meta-data header: " + s);
								}
								else
								{
									string Key = s[..j].Trim();
									string Value = s[(j + 1)..].Trim();

									Headers ??= new List<KeyValuePair<string, string>>();
									Headers.Add(new KeyValuePair<string, string>(Key, Value));
								}
							}
							else
							{
								Error = true;
								Console.Error.WriteLine("Missing meta-data header.");
							}
							break;

						case "-r":
						case "-recursive":
							Recursive = true;
							break;

						case "-?":
						case "-h":
						case "-help":
							Console.Out.WriteLine("This tool converts a Word document to a Markdown document.");
							Console.Out.WriteLine();
							Console.Out.WriteLine("Syntax: WordToMarkdown -input WORD_FILENAME -output MARKDOWN_FILENAME");
							Console.Out.WriteLine();
							Console.Out.WriteLine("Following switches are recognized:");
							Console.Out.WriteLine();
							Console.Out.WriteLine("-i FILENAME        Defines the filename of the Word document. The");
							Console.Out.WriteLine("-input FILENAME    Word document must be saved using the Open XML");
							Console.Out.WriteLine("-word FILENAME     SDK (i.e. in .docx file format). Filename can");
							Console.Out.WriteLine("                   contain wildcards (*).");
							Console.Out.WriteLine("-o FILENAME        Defines the filename of the Markdown document");
							Console.Out.WriteLine("-output FILENAME   that will be generated. This switch is optional.");
							Console.Out.WriteLine("-md FILENAME       If not provided, the same file name as the Word");
							Console.Out.WriteLine("-markdown FILENAME document will be used, with the file extension .md.");
							Console.Out.WriteLine("                   Filename can contain wildcards matching input");
							Console.Out.WriteLine("                   filename.");
							Console.Out.WriteLine("-meta KEY=VALUE    Adds a Markdown header to the output. Reference:");
							Console.Out.WriteLine("-header KEY=VALUE  https://lab.tagroot.io/Markdown.md#metadata");
							Console.Out.WriteLine("-r                 Recursive search for documents.");
							Console.Out.WriteLine("-recursive         Same as -r.");
							Console.Out.WriteLine("-?, -h, -help:     Shows this help.");
							break;

						default:
							Error = true;
							Console.Error.WriteLine("Unrecognized switch: " + e.Args[i - 1]);
							break;
					}

					if (Error)
						break;
				}

				if (!Error)
				{
					if (string.IsNullOrEmpty(InputFileName))
					{
						Error = true;
						Console.Error.WriteLine("Missing input file name.");
					}
					else if (!string.IsNullOrEmpty(OutputFileName) &&
						OutputFileName.Contains('*') &&
						InputFileName.Split('*').Length != OutputFileName.Split('*').Length)
					{
						Error = true;
						Console.Error.WriteLine("Number of wildcards do not match.");
					}
					else
					{
						try
						{
							KeyValuePair<int, int> P = ConvertWithWildcard(InputFileName, OutputFileName, Recursive, Headers?.ToArray());
							int NrConverted = P.Key;
							int NrNotConverted = P.Value;

							Console.Out.WriteLine(GetMessage(NrConverted, NrNotConverted));
						}
						catch (Exception ex)
						{
							Console.Out.WriteLine(ex.Message);
							Error = true;
						}
					}
				}

				this.Shutdown(Error ? 1 : 0);
			}
		}

		public static string GetMessage(int NrConverted, int NrNotConverted)
		{
			if (NrConverted == 0)
			{
				if (NrNotConverted == 0)
					return "No files processed.";
				else if (NrNotConverted == 1)
					return "File not converted.";
				else
					return NrNotConverted.ToString() + " files found, bot not converted.";
			}
			else if (NrNotConverted == 0)
			{
				if (NrConverted == 0)
					return "No files processed.";
				else if (NrConverted == 1)
					return "File converted.";
				else
					return NrConverted.ToString() + " files converted.";
			}
			else
			{
				int NrFiles = NrConverted + NrNotConverted;

				return NrFiles.ToString() + " file(s) found, " +
					NrConverted.ToString() + " where converted, " +
					NrNotConverted.ToString() + " where not converted.";
			}
		}

		/// <summary>
		/// Converts a collection of Word documents to Markdown.
		/// </summary>
		/// <param name="InputFileName">Input file name, possibly containing wildcards (*).</param>
		/// <param name="OutputFileName">optional Output file name, possibly containing wildcards (the same amount as for the input file name).</param>
		/// <param name="Recursive">If search should be recursive.</param>
		/// <param name="Headers">Additional headers to add to Markdown output.</param>
		/// <returns>Number of files converted, number of files not converted.</returns>
		public static KeyValuePair<int, int> ConvertWithWildcard(string InputFileName, string? OutputFileName, bool Recursive,
			params KeyValuePair<string, string>[]? Headers)
		{
			string? Folder = Path.GetDirectoryName(InputFileName);
			if (string.IsNullOrEmpty(Folder))
				Folder = Environment.CurrentDirectory;
			else if (Folder.Contains('*'))
				throw new Exception("Folder cannot contain wildcards.");

			string FileName = Path.GetFileName(InputFileName);
			string[] Parts = FileName.Split('*', StringSplitOptions.None);
			int NrConverted = 0;
			int NrNotConverted = 0;

			if (Parts.Length == 1 && !Recursive)
			{
				if (ConvertIndividualFile(InputFileName, OutputFileName, Headers))
					NrConverted++;
				else
					NrNotConverted++;
			}
			else
			{
				StringBuilder RegexBuilder = new();
				bool First = true;
				int NrParameters = 0;
				int i, j, c;

				RegexBuilder.Append('^');

				foreach (string Part in Parts)
				{
					if (First)
						First = false;
					else
					{
						RegexBuilder.Append("(?'P");
						RegexBuilder.Append(NrParameters++);
						RegexBuilder.Append("'.*)");
					}

					i = 0;
					c = Part.Length;
					while (i < c)
					{
						j = Part.IndexOfAny(regexSpecialCharaters, i);
						if (j < i)
						{
							RegexBuilder.Append(Part[i..]);
							i = c;
						}
						else
						{
							if (j > i)
								RegexBuilder.Append(Part[i..j]);

							RegexBuilder.Append('\\');
							RegexBuilder.Append(Part[j]);

							i = j + 1;
						}
					}
				}

				RegexBuilder.Append('$');

				Regex Parsed = new(RegexBuilder.ToString(), RegexOptions.Singleline | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
				string[] Files = Directory.GetFiles(Folder, FileName, Recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly);

				foreach (string File in Files)
				{
					if (string.IsNullOrEmpty(OutputFileName))
					{
						if (ConvertIndividualFile(File, null, Headers))
							NrConverted++;
						else
							NrNotConverted++;
					}
					else
					{
						Match M = Parsed.Match(Path.GetFileName(File));
						if (!M.Success)
							continue;

						string s = OutputFileName;
						string s2;

						for (i = j = 0; i < NrParameters; i++)
						{
							j = s.IndexOf('*', j);
							if (j < 0)
								break;

							s2 = M.Groups["P" + i.ToString()].Value;
							s = s.Remove(j, 1).Insert(j, s2);
							j += s2.Length;
						}

						if (ConvertIndividualFile(File, s, Headers))
							NrConverted++;
						else
							NrNotConverted++;
					}
				}
			}

			return new KeyValuePair<int, int>(NrConverted, NrNotConverted);
		}

		private static readonly char[] regexSpecialCharaters = new char[] { '\\', '^', '$', '{', '}', '[', ']', '(', ')', '.', '*', '+', '?', '|', '<', '>', '-', '&' };
		private static readonly UTF8Encoding utf8Bom = new(true);

		/// <summary>
		/// Converts a Word file to markdown.
		/// </summary>
		/// <param name="InputFileName">Name of Word file.</param>
		/// <param name="OutputFileName">Optional name of Output file.</param>
		/// <param name="Headers">Additional headers to add to Markdown output.</param>
		/// <returns>If conversion was possible.</returns>
		public static bool ConvertIndividualFile(string InputFileName, string? OutputFileName,
			params KeyValuePair<string, string>[]? Headers)
		{
			if (string.IsNullOrEmpty(OutputFileName))
				OutputFileName = Path.ChangeExtension(InputFileName, "md");
			else if (Directory.Exists(OutputFileName))
			{
				string FileName = Path.GetFileName(InputFileName);
				FileName = Path.ChangeExtension(FileName, "md");

				if (!OutputFileName.EndsWith(Path.DirectorySeparatorChar))
					OutputFileName += Path.DirectorySeparatorChar;

				OutputFileName = Path.Combine(OutputFileName, FileName);
			}

			Console.Out.WriteLine("Processing: " + InputFileName);

			WordprocessingDocument? Doc = null;
			string? TempFileName = null;

			try
			{
				try
				{
					try
					{
						Doc = WordprocessingDocument.Open(InputFileName, false);
					}
					catch (OpenXmlPackageException)
					{
						TempFileName = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
						File.Copy(InputFileName, TempFileName);

						Doc = WordprocessingDocument.Open(TempFileName, true, WordUtilities.GetFailSafePackageSettings());
					}
				}
				catch (Exception)
				{
					Console.Out.WriteLine("Unable to open Word file: " + InputFileName);
					return false;
				}

				string Markdown = WordUtilities.ExtractAsMarkdown(Doc, InputFileName, out _);

				Dictionary<string, bool> HeadersUsed = new();
				StringBuilder sb = new();
				DateTime? TP;
				string? s;
				bool HeadersAdded = false;

				if (Headers is not null)
				{
					foreach (KeyValuePair<string, string> Header in Headers)
					{
						HeadersUsed[Header.Key] = true;
						AppendHeader(sb, Header.Key, Header.Value, ref HeadersAdded);
					}
				}

				if (!string.IsNullOrEmpty(s = Doc.PackageProperties.Category) && !HeadersUsed.ContainsKey("Cagegory"))
					AppendHeader(sb, "Category", s, ref HeadersAdded);

				if (!string.IsNullOrEmpty(s = Doc.PackageProperties.Language) && !HeadersUsed.ContainsKey("Language"))
					AppendHeader(sb, "Language", s, ref HeadersAdded);

				if (!string.IsNullOrEmpty(s = Doc.PackageProperties.Version) && !HeadersUsed.ContainsKey("Version"))
					AppendHeader(sb, "Version", s, ref HeadersAdded);

				if (!string.IsNullOrEmpty(s = Doc.PackageProperties.Title) && !HeadersUsed.ContainsKey("Title"))
					AppendHeader(sb, "Title", s, ref HeadersAdded);

				if (!string.IsNullOrEmpty(s = Doc.PackageProperties.Subject) && !HeadersUsed.ContainsKey("Subject"))
					AppendHeader(sb, "Subject", s, ref HeadersAdded);

				if (!string.IsNullOrEmpty(s = Doc.PackageProperties.Creator) && !HeadersUsed.ContainsKey("Author"))
					AppendHeader(sb, "Author", s, ref HeadersAdded);

				if (!string.IsNullOrEmpty(s = Doc.PackageProperties.Keywords) && !HeadersUsed.ContainsKey("Keywords"))
					AppendHeader(sb, "Keywords", s, ref HeadersAdded);

				if (!string.IsNullOrEmpty(s = Doc.PackageProperties.Description) && !HeadersUsed.ContainsKey("Description"))
					AppendHeader(sb, "Description", s, ref HeadersAdded);

				if ((TP = Doc.PackageProperties.Created).HasValue && !HeadersUsed.ContainsKey("Date"))
					AppendHeader(sb, "Date", TP.Value.Date.ToShortDateString(), ref HeadersAdded);

				if (HeadersAdded)
				{
					sb.AppendLine();
					sb.Append(Markdown);
					Markdown = sb.ToString();
				}

				Console.Out.WriteLine("Saving: " + OutputFileName);

				File.WriteAllText(OutputFileName, Markdown, utf8Bom);

				return true;
			}
			finally
			{
				Doc?.Dispose();

				if (!string.IsNullOrEmpty(TempFileName) && File.Exists(TempFileName))
					File.Delete(TempFileName);
			}
		}

		private static void AppendHeader(StringBuilder Markdown, string Key, string Value, ref bool HeadersAdded)
		{
			foreach (string Row in Value.Trim().Replace("\r\n", "\n").Replace('\r', '\n').Split('\n', StringSplitOptions.RemoveEmptyEntries))
			{
				Markdown.Append(Key);
				Markdown.Append(": ");
				Markdown.AppendLine(Row);
				HeadersAdded = true;
			}
		}
	}
}
