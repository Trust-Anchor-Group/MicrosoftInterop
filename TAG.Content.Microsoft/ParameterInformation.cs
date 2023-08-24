using System.Collections.Generic;

namespace TAG.Content.Microsoft
{
	/// <summary>
	/// Information about a parameter found in a Markdown document generated from a Word document.
	/// </summary>
	public class ParameterInformation
	{
		/// <summary>
		/// Parameter Values found.
		/// </summary>
		public List<string> Values = new List<string>();

		/// <summary>
		/// Other properties found relating to the paramter.
		/// </summary>
		public Dictionary<string, List<string>> Properties = new Dictionary<string, List<string>>();

		/// <summary>
		/// Options associated with parameter.
		/// </summary>
		public List<OptionInformation> Options = null;

		/// <summary>
		/// Name of parameter.
		/// </summary>
		public string Name;

		/// <summary>
		/// Type of parameter
		/// </summary>
		public ParameterType Type;

		/// <summary>
		/// Any maximum length information available.
		/// </summary>
		public int? MaxLength;
	}
}
