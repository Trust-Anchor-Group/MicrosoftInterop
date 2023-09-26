using System;
using System.Collections.Generic;
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
						InputFileName.Split('*').Length != OutputFileName.Split('*').Length)
					{
						Error = true;
						Console.Error.WriteLine("Number of wildcards do not match.");
					}
					else
					{
						try
						{
							ConvertWithWildcard(InputFileName, OutputFileName, Recursive, Headers?.ToArray());
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

		/// <summary>
		/// Converts a collection of Word documents to Markdown.
		/// </summary>
		/// <param name="InputFileName">Input file name, possibly containing wildcards (*).</param>
		/// <param name="OutputFileName">optional Output file name, possibly containing wildcards (the same amount as for the input file name).</param>
		/// <param name="Recursive">If search should be recursive.</param>
		/// <param name="Headers">Additional headers to add to Markdown output.</param>
		public static void ConvertWithWildcard(string InputFileName, string? OutputFileName, bool Recursive,
			params KeyValuePair<string, string>[]? Headers)
		{
			string? Folder = Path.GetDirectoryName(InputFileName);
			if (string.IsNullOrEmpty(Folder))
				Folder = Environment.CurrentDirectory;
			else if (Folder.Contains('*'))
				throw new Exception("Folder cannot contain wildcards.");

			string FileName = Path.GetFileName(InputFileName);
			string[] Parts = FileName.Split('*', StringSplitOptions.None);
			if (Parts.Length == 1 && !Recursive)
			{
				ConvertIndividualFile(InputFileName, OutputFileName, Headers);
				return;
			}

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

			Regex Parsed = new(RegexBuilder.ToString(), RegexOptions.Singleline | RegexOptions.CultureInvariant);
			string[] Files = Directory.GetFiles(Folder, FileName, Recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly);

			foreach (string File in Files)
			{
				if (string.IsNullOrEmpty(OutputFileName))
					ConvertIndividualFile(File, null, Headers);
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

					ConvertIndividualFile(File, s, Headers);
				}
			}
		}

		private static readonly char[] regexSpecialCharaters = new char[] { '\\', '^', '$', '{', '}', '[', ']', '(', ')', '.', '*', '+', '?', '|', '<', '>', '-', '&' };
		private static readonly UTF8Encoding utf8Bom = new(true);

		/// <summary>
		/// Converts a Word file to markdown.
		/// </summary>
		/// <param name="InputFileName">Name of Word file.</param>
		/// <param name="OutputFileName">Optional name of Output file.</param>
		/// <param name="Headers">Additional headers to add to Markdown output.</param>
		public static void ConvertIndividualFile(string InputFileName, string? OutputFileName,
			params KeyValuePair<string, string>[]? Headers)
		{
			if (string.IsNullOrEmpty(OutputFileName))
				OutputFileName = Path.ChangeExtension(InputFileName, "md");

			Console.Out.WriteLine("Processing: " + InputFileName);

			string Markdown = WordUtilities.ExtractAsMarkdown(InputFileName);

			if (Headers is not null && Headers.Length > 0)
			{
				int? HeaderEndPos = MarkdownDocument.HeaderEndPosition(Markdown);
				if (!HeaderEndPos.HasValue)
				{
					Markdown = "\r\n\r\n" + Markdown;
					HeaderEndPos = 0;
				}

				StringBuilder sb = new();

				foreach (KeyValuePair<string, string> Header in Headers)
				{
					sb.Append(Header.Key);
					sb.Append(": ");
					sb.AppendLine(Header.Value);
				}

				Markdown = Markdown.Insert(HeaderEndPos.Value, sb.ToString());
			}

			Console.Out.WriteLine("Saving: " + OutputFileName);

			File.WriteAllText(OutputFileName, Markdown, utf8Bom);
		}
	}
}
