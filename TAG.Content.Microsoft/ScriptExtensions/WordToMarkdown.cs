using DocumentFormat.OpenXml.Packaging;
using Waher.Script;
using Waher.Script.Abstraction.Elements;
using Waher.Script.Exceptions;
using Waher.Script.Model;
using Waher.Script.Objects;

namespace TAG.Content.Microsoft.ScriptExtensions
{
	/// <summary>
	/// Converts a Word document to Markdown.
	/// </summary>
	public class WordToMarkdown : FunctionOneScalarVariable
	{
		/// <summary>
		/// Converts a Word document to Markdown.
		/// </summary>
		/// <param name="Doc">Word document</param>
		/// <param name="Start">Start position in the underlying script.</param>
		/// <param name="Length">Length of element in underlying script.</param>
		/// <param name="Expression">Expression object.</param>
		public WordToMarkdown(ScriptNode Doc, int Start, int Length, Expression Expression)
			: base(Doc, Start, Length, Expression)
		{
		}

		/// <summary>
		/// Name of function
		/// </summary>
		public override string FunctionName => nameof(WordToMarkdown);

		/// <summary>
		/// Evaluates the function.
		/// </summary>
		/// <param name="Argument">Evaluated argument.</param>
		/// <param name="Variables">Variables collection.</param>
		/// <returns>Result</returns>
		public override IElement EvaluateScalar(IElement Argument, Variables Variables)
		{
			if (!(Argument.AssociatedObjectValue is WordprocessingDocument Doc))
				throw new ScriptRuntimeException("Expected a Word document.", this);

			string Markdown = WordUtilities.ExtractAsMarkdown(Doc);

			return new StringValue(Markdown);
		}
	}
}
