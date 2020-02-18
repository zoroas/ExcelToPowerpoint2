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

namespace ExcelToPowerpoint2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = this;
            this.ExcelFileName = MySettings.Default.SettingExcelFileName;
            this.PowerPointTemplateFileName = MySettings.Default.SettingPowerPointTemplateFileName;
            this.PowerPointOutputFileName = MySettings.Default.SettingPowerPointOutputFileName;
            this.PhotoFolderName = MySettings.Default.SettingPhotoFolderName;
            this.NumberOfColumns = MySettings.Default.SettingNumberOfColumns;
            this.RowNumberToGetColumns = MySettings.Default.SettingRowNumberToGetColumns;
        }

        public static readonly DependencyProperty ExcelFileNameProperty =
            DependencyProperty.Register("ExcelFileName",
                typeof(string), typeof(MainWindow), new PropertyMetadata(""));
        
        public static readonly DependencyProperty PowerPointTemplateFileNameProperty =
            DependencyProperty.Register("PowerPointTemplateFileName",
                typeof(string), typeof(MainWindow), new PropertyMetadata(""));

        public static readonly DependencyProperty NumberOfColumnsProperty =
            DependencyProperty.Register("NumberOfColumns",
                typeof(string), typeof(MainWindow), new PropertyMetadata(""));

        public static readonly DependencyProperty RowNumberToGetColumnsProperty =
            DependencyProperty.Register("RowNumberToGetColumns",
                 typeof(string), typeof(MainWindow), new PropertyMetadata(""));

        public static readonly DependencyProperty PowerPointOutputFileNameProperty =
            DependencyProperty.Register("PowerPointOutputFileName",
                typeof(string), typeof(MainWindow), new PropertyMetadata(""));

        public static readonly DependencyProperty PhotoFolderNameProperty =
            DependencyProperty.Register("PhotoFolderName",
                typeof(string), typeof(MainWindow), new PropertyMetadata(""));

        public string ExcelFileName
        {
            get { return (string)GetValue(ExcelFileNameProperty); }
            set
            {
                SetValue(ExcelFileNameProperty, value);
                MySettings.Default.SettingExcelFileName = value;
            }
        }

        public string PowerPointTemplateFileName
        {
            get { return (string)GetValue(PowerPointTemplateFileNameProperty); }
            set
            {
                SetValue(PowerPointTemplateFileNameProperty, value);
                MySettings.Default.SettingPowerPointTemplateFileName = value;
            }
        }

        public string PowerPointOutputFileName
        {
            get { return (string)GetValue(PowerPointOutputFileNameProperty); }
            set
            {
                SetValue(PowerPointOutputFileNameProperty, value);
                MySettings.Default.SettingPowerPointOutputFileName = value;
            }
        }

        public string RowNumberToGetColumns
        {
            get { return (string)GetValue(RowNumberToGetColumnsProperty); }
            set
            {
                SetValue(RowNumberToGetColumnsProperty, value);
            }
        }

        public string NumberOfColumns
        {
            get { return (string)GetValue(NumberOfColumnsProperty); }
            set
            {
                SetValue(NumberOfColumnsProperty, value);
            }
        }

        public string PhotoFolderName
        {
            get { return (string)GetValue(PhotoFolderNameProperty); }
            set
            {
                SetValue(PhotoFolderNameProperty, value);
            }
        }

        private void TBExcelDataFile_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void TBPowerPoint_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void TBPhotoFolder_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void TBPhotoFolder_Drop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    this.PhotoFolderName = files[0];
                }
            }
            catch (Exception) { }
        }

        private void ButtonPhotoFolder_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog 
            System.Windows.Forms.FolderBrowserDialog dlg = new System.Windows.Forms.FolderBrowserDialog();

            dlg.ShowDialog();
            // Get the selected file name and display in a TextBox 
            if (dlg.SelectedPath != null)
            {
                // Open document 
                string folder = dlg.SelectedPath;
                this.PhotoFolderName = folder;
            }
        }

        private void TBExcelDataFile_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                this.ExcelFileName = files[0];
            }
        }

        private void TBPowerPoint_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                this.PowerPointTemplateFileName = files[0];
            }
        }


        private void TBOutputPowerPoint_Drop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    this.PowerPointOutputFileName = files[0];
                }
            }
            catch (Exception) { }
        }

        private void TBOutputPowerPoint_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void ButtonExcelDataFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel documents (.xlsx)|*.xlsx|Excel old documents (.xls)|*.xls|All files (*.*)|*.*";

            // Get the selected file name and display in a TextBox 
            if (dlg.ShowDialog() == true)
            {
                // Open document 
                string filename = dlg.FileName;
                this.ExcelFileName = filename;
            }
        }

        private void ButtonPowerPointTemplate_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".pptx";
            dlg.Filter = "PowerPoint documents (.pptx)|*.pptx|PowerPoint old documents (.ppt)|*.ppt|All files (*.*)|*.*";

            // Get the selected file name and display in a TextBox 
            if (dlg.ShowDialog() == true)
            {
                // Open document 
                string filename = dlg.FileName;
                this.PowerPointTemplateFileName = filename;
            }
        }

        private void ButtonOutputPowerPoint_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".pptx";
            dlg.Filter = "PowerPoint documents (.pptx)|*.pptx|PowerPoint old documents (.ppt)|*.ppt|All files (*.*)|*.*";

            // Get the selected file name and display in a TextBox 
            if (dlg.ShowDialog() == true)
            {
                // Open document 
                string filename = dlg.FileName;
                this.PowerPointOutputFileName = filename;
            }
        }

        private void ButtonOptions_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ButtonExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void ButtonCreateFile_Click(object sender, RoutedEventArgs e)
        {
            Cursor = Cursors.Wait;

            PowerPointCreator c = new PowerPointCreator(this.ExcelFileName, this.PowerPointTemplateFileName,
                this.PowerPointOutputFileName, this.RowNumberToGetColumns,
                this.NumberOfColumns, this.PhotoFolderName);
            if (c.HasError())
                MessageBox.Show(c.GetError());
            Cursor = Cursors.Arrow;
        }
    }
}
