using System.Diagnostics;
using System.Windows;
using System.Windows.Documents;

namespace WordToMarkdown
{
	/// <summary>
	/// Interaction logic for HelpWindow.xaml
	/// </summary>
	public partial class HelpWindow : Window
	{
		public HelpWindow()
		{
			InitializeComponent();
		}

		private void CloseWindow(object sender, RoutedEventArgs e)
		{
			this.Close();
		}

		private void Hyperlink_Click(object sender, RoutedEventArgs e)
		{
			if (sender is Hyperlink Link)
			{
				ProcessStartInfo Info = new ProcessStartInfo()
				{
					FileName = Link.NavigateUri.ToString(),
					UseShellExecute = true
				};

				Process.Start(Info);
			}
		}
	}
}
