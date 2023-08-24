using System;
using System.Collections.Generic;
using Waher.Content.Markdown;

namespace TAG.Content.Microsoft
{
	/// <summary>
	/// Utilities for working with Markdown from Microsoft Office Word documents, in Smart Contracts.
	/// </summary>
	public static class ContractUtilities
	{
		/// <summary>
		/// Extract parameters from a Markdown document, earlier identified by a conversion of a Microsoft Word document,
		/// and exported to Markdown. Such parameters are encoded into the header of the document. If such a header is
		/// found, it is also removed from the Markdown.
		/// </summary>
		/// <param name="Markdown">Markdown document.</param>
		/// <param name="ByName">Parameter information, by name.</param>
		/// <returns>If parameters were found in the Markdown document.</returns>
		public static bool ExtractParameters(ref string Markdown, out Dictionary<string, ParameterInformation> ByName)
		{
			int? Pos = MarkdownDocument.HeaderEndPosition(Markdown);
			if (!Pos.HasValue)
			{
				ByName = null;
				return false;
			}

			string Header = Markdown.Substring(0, Pos.Value);
			List<string> Values;

			Markdown = Markdown.Substring(Pos.Value).TrimStart();

			string[] Rows = Header.
				Replace("\r\n", "\n").
				Replace('\r', '\n').
				Split('\n');

			ByName = new Dictionary<string, ParameterInformation>();

			foreach (string Row in Rows)
			{
				int i = Row.IndexOf(':');
				if (i > 0)
				{
					string Key = Row.Substring(0, i).Trim();
					string Value = Row.Substring(i + 1).Trim();

					if (ByName.TryGetValue(Key, out ParameterInformation Info))
						Values = Info.Values;
					else
					{
						Values = null;
						bool Handled = false;

						foreach (KeyValuePair<string, ParameterInformation> P in ByName)
						{
							if (Key.StartsWith(P.Key + " "))
							{
								string s = Key.Substring(P.Key.Length + 1).TrimStart();

								if (s == "Type")
								{
									if (Enum.TryParse(Value, out ParameterType ParameterType))
										P.Value.Type = ParameterType;
									else
										P.Value.Type = ParameterType.String;

									Handled = true;
								}
								else if (s == "MaxLen")
								{
									if (int.TryParse(Value, out int MaxLength))
										P.Value.MaxLength = MaxLength;
									else
										P.Value.MaxLength = null;

									Handled = true;
								}
								else if (s.StartsWith("Item") &&
									s.EndsWith(" Value") &&
									int.TryParse(s.Substring(4, s.Length - 10), out int ItemIndex))
								{
									if (P.Value.Options is null)
										P.Value.Options = new List<OptionInformation>();

									while (P.Value.Options.Count < ItemIndex)
										P.Value.Options.Add(new OptionInformation());

									P.Value.Options[ItemIndex - 1].Value = Value;
									Handled = true;
								}
								else if (s.StartsWith("Item") &&
									s.EndsWith(" Display") &&
									int.TryParse(s.Substring(4, s.Length - 12), out ItemIndex))
								{
									if (P.Value.Options is null)
										P.Value.Options = new List<OptionInformation>();

									while (P.Value.Options.Count < ItemIndex)
										P.Value.Options.Add(new OptionInformation());

									P.Value.Options[ItemIndex - 1].Display = Value;
									Handled = true;
								}
								else if (!P.Value.Properties.TryGetValue(s, out Values))
								{
									Values = new List<string>();
									P.Value.Properties[s] = Values;
								}

								break;
							}
						}

						if (Handled)
							continue;

						if (Values is null)
						{
							Info = new ParameterInformation()
							{
								Name = Key
							};

							ByName[Key] = Info;
							Values = Info.Values;
						}
					}

					Values.Add(Value);
				}
			}

			return ByName.Count > 0;
		}
	}
}
