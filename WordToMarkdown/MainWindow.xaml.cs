using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;
using Waher.Events;

namespace WordToMarkdown
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window, INotifyPropertyChanged
	{
		private string inputFileName = string.Empty;
		private string outputFileName = string.Empty;

		public MainWindow()
		{
			this.InitializeComponent();
			this.DataContext = this;
		}

		public string InputFileName
		{
			get => this.inputFileName;
			set
			{
				if (this.inputFileName != value)
				{
					this.inputFileName = value;
					this.RaisePropertyChanged(nameof(this.InputFileName));
				}
			}
		}

		public string OutputFileName
		{
			get => this.outputFileName;
			set
			{
				if (this.outputFileName != value)
				{
					this.outputFileName = value;
					this.RaisePropertyChanged(nameof(this.OutputFileName));
				}
			}
		}

		public bool Recursive { get; set; } = false;

		public event PropertyChangedEventHandler? PropertyChanged;

		private void RaisePropertyChanged(string Name)
		{
			this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(Name));
		}

		private void CloseButtonClicked(object sender, RoutedEventArgs e)
		{
			this.Close();
		}

		private void ConvertButtonClicked(object sender, RoutedEventArgs e)
		{
			Mouse.OverrideCursor = Cursors.Wait;
			try
			{
				KeyValuePair<int, int> P = App.ConvertWithWildcard(this.InputFileName, this.OutputFileName, this.Recursive);
				int NrConverted = P.Key;
				int NrNotConverted = P.Value;
				string Message = App.GetMessage(NrConverted, NrNotConverted);

				Mouse.OverrideCursor = null;

				if (NrNotConverted>0)
					MessageBox.Show(this, Message, "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
				else
					MessageBox.Show(this, Message, "Information", MessageBoxButton.OK, MessageBoxImage.Information);
			}
			catch (Exception ex)
			{
				Mouse.OverrideCursor = null;
				ex = Log.UnnestException(ex);
				MessageBox.Show(this, ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void BrowseWordDocuments(object sender, RoutedEventArgs e)
		{
			OpenFileDialog Dialog = new()
			{
				DefaultExt = "*.docx",
				Filter = "Word documents (*.docx)|*.docx",
				CheckFileExists = true,
				CheckPathExists = true,
				Multiselect = false,
				ShowReadOnly = true,
				Title = "Select Word Document"
			};

			bool? Result = Dialog.ShowDialog();

			if (Result.HasValue && Result.Value)
				this.InputFileName = Dialog.FileName;
		}

		private void BrowseMarkdownFiles(object sender, RoutedEventArgs e)
		{
			OpenFileDialog Dialog = new()
			{
				DefaultExt = "*.md",
				Filter = "Markdown files (*.md)|*.md",
				CheckFileExists = false,
				CheckPathExists = true,
				Multiselect = false,
				ShowReadOnly = true,
				Title = "Select Markdown File"
			};

			bool? Result = Dialog.ShowDialog();

			if (Result.HasValue && Result.Value)
				this.OutputFileName = Dialog.FileName;
		}
	}
}
