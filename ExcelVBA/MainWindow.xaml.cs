using System;
using System.Windows;
using System.IO;
using System.Windows.Forms;

namespace ExcelVBA
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        
       


        private void button_Click(object sender, RoutedEventArgs e)
        {
           

                
                Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog();
                ofd.Multiselect = true;
                ofd.DefaultExt = ".xls";
                ofd.Filter = ".xls|*.xls";
                Nullable<bool> dialogOK = ofd.ShowDialog();
                string glownyplik = ofd.InitialDirectory + ofd.FileName;
                FolderBrowserDialog katalog = new FolderBrowserDialog();
                katalog.ShowDialog();
                string sourceDirectory = katalog.SelectedPath;
                textBox.Text = sourceDirectory;
                var txtFiles = Directory.EnumerateFiles(sourceDirectory, "*.xls*", SearchOption.AllDirectories);
                CopyModule kopia = new CopyModule();
                       
                foreach (string currentFile in txtFiles)
                {
                    kopia.CopyMacro(glownyplik, currentFile);
                }
            GC.Collect();
           

        }
     
        private void textBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            
        }
    }
}


       
    



