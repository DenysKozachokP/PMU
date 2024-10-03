using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Controls.Ribbon;
using Microsoft.Win32;
using System.IO;

namespace KWord
{
    public partial class MainWindow : RibbonWindow
    {
        public bool IsSaved = false;
        private void doc1_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }

    public partial class MainWindow : RibbonWindow
    {
        public ICommand ChangeColorCommand { get; set; }
        public MainWindow()
        {
            InitializeComponent();
            _fontSize.ItemsSource = FontSizes;

            ChangeColorCommand = new RelayCommand<string>(ChangeTextColor);
            this.DataContext = this;
        }
        private void ChangeTextColor(string colorName)
        {
            TextSelection selectedText = doc1.Selection;
            if (!selectedText.IsEmpty)
            {
                selectedText.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush((Color)ColorConverter.ConvertFromString(colorName)));
            }
        }
        public double[] FontSizes
        {
            get
            {
                return new double[] { 3.0, 4.0, 5.0, 6.0, 6.5, 7.0, 7.5, 8.0, 8.5, 9.0,
                9.5, 10.0, 10.5, 11.0, 11.5, 12.0, 12.5,13.0,13.5,14.0, 15.0,16.0, 17.0, 18.0, 19.0,
                20.0, 22.0, 24.0, 26.0, 28.0, 30.0,32.0, 34.0, 36.0, 38.0, 40.0, 44.0, 48.0, 52.0, 56.0,
                60.0, 64.0, 68.0, 72.0, 76.0,80.0, 88.0, 96.0, 104.0, 112.0, 120.0, 128.0, 136.0, 144.0
                };
            }
        }
        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog(); dlg.Filter = "Document files (*.rtf)|*.rtf"; var result = dlg.ShowDialog();
            if (result.Value)
            {
                TextRange t = new TextRange(doc1.Document.ContentStart, doc1.Document.ContentEnd); FileStream file = new FileStream(dlg.FileName, FileMode.Open);
                t.Load(file, System.Windows.DataFormats.Rtf);
            }
        }
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog savefile = new SaveFileDialog();
            savefile.FileName = "unknown.doc";
            savefile.Filter = "Document files (*.rtf)|*.rtf";
            if (savefile.ShowDialog() == true)
            {
                TextRange t = new TextRange(doc1.Document.ContentStart, doc1.Document.ContentEnd); this.Title = this.Title + " " + savefile.FileName;
                FileStream file = new FileStream(savefile.FileName, FileMode.Create); t.Save(file, System.Windows.DataFormats.Rtf);
                file.Close();
            }
            IsSaved = true;
        }
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            if (IsSaved == false)
                if (MessageBox.Show("Do you want save changes ?", "Message", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    this.btnSave_Click(sender, e);
                }
            this.Close();
        }
        void ApplyPropertyValueToSelectedText(DependencyProperty formattingProperty, object value)
        {
            if (value == null) return;
            doc1.Selection.ApplyPropertyValue(formattingProperty, value);
        }
        private void FontFamili_SelectionChange(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                FontFamily editValue = (FontFamily)e.AddedItems[0]; ApplyPropertyValueToSelectedText(TextElement.FontFamilyProperty, editValue);
            }
            catch (Exception) { }
        }
        private void FontSize_SelectionChange(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                ApplyPropertyValueToSelectedText(TextElement.FontSizeProperty, e.AddedItems[0]);
            }
            catch (Exception) { }
        }
        private void EditHeader_Click(object sender, RoutedEventArgs e)
        {
            HeaderTextBox.Focus();
        }

        private void EditFooter_Click(object sender, RoutedEventArgs e)
        {
            FooterTextBox.Focus();
        }
        private void HelpButton_Click(object sender, RoutedEventArgs e)
        {
            HelpWindow helpWindow = new HelpWindow();
            helpWindow.ShowDialog();
        }
    }
}
