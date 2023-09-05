using DocumentFormat.OpenXml.Packaging;
using Waher.Script;
using Waher.Script.Abstraction.Elements;
using Waher.Script.Exceptions;
using Waher.Script.Model;
using Waher.Script.Objects;

namespace TAG.Content.Microsoft.ScriptExtensions
{
	/// <summary>
	/// Converts a Excel document to Script.
	/// </summary>
	public class ExcelToScript : FunctionOneScalarVariable
	{
		/// <summary>
		/// Converts a Excel document to Script.
		/// </summary>
		/// <param name="Doc">Excel document</param>
		/// <param name="Start">Start position in the underlying script.</param>
		/// <param name="Length">Length of element in underlying script.</param>
		/// <param name="Expression">Expression object.</param>
		public ExcelToScript(ScriptNode Doc, int Start, int Length, Expression Expression)
			: base(Doc, Start, Length, Expression)
		{
		}

		/// <summary>
		/// Name of function
		/// </summary>
		public override string FunctionName => nameof(ExcelToScript);

		/// <summary>
		/// Evaluates the function.
		/// </summary>
		/// <param name="Argument">Evaluated argument.</param>
		/// <param name="Variables">Variables collection.</param>
		/// <returns>Result</returns>
		public override IElement EvaluateScalar(IElement Argument, Variables Variables)
		{
			if (!(Argument.AssociatedObjectValue is SpreadsheetDocument Doc))
				throw new ScriptRuntimeException("Expected an Excel document.", this);

			string Script = ExcelUtilities.ExtractAsScript(Doc, true);

			return new StringValue(Script);
		}
	}
}
