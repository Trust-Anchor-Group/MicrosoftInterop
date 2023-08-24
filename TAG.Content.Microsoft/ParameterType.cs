namespace TAG.Content.Microsoft
{
	/// <summary>
	/// Type of parameter.
	/// </summary>
	public enum ParameterType
	{
		/// <summary>
		/// String input.
		/// </summary>
		String,

		/// <summary>
		/// Checkbox
		/// </summary>
		Boolean,

		/// <summary>
		/// ComboBox with drop-down list.
		/// </summary>
		StringWithOptions,

		/// <summary>
		/// List-box
		/// </summary>
		ListOfOptions,

		/// <summary>
		/// Date picker.
		/// </summary>
		Date,

		/// <summary>
		/// Time entry.
		/// </summary>
		Time,

		/// <summary>
		/// Number entry.
		/// </summary>
		Number
	}
}
