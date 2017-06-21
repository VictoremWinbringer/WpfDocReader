using System;
using System.Windows;
using System.IO;
using Microsoft.Win32;
using Spire.Doc;
using System.Windows.Xps.Packaging;

namespace WpfApp2
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

        private void btnSelectWord_Click(object sender, RoutedEventArgs e)
        {
            // Initialize an OpenFileDialog 
            OpenFileDialog openFileDialog = new OpenFileDialog();


            // Set filter and RestoreDirectory 
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Filter = "Word documents(*.doc;*.docx)|*.doc;*.docx";


            bool? result = openFileDialog.ShowDialog();
            if (result == true)
            {
                if (openFileDialog.FileName.Length > 0)
                {
                    txbSelectedWordFile.Text = openFileDialog.FileName;
                }
            }

        }

        private XpsDocument ConvertWordToXps(string wordFilename, string xpsFilename)
        {
            Document doc = new Document(wordFilename);

            doc.SaveToFile(xpsFilename, FileFormat.XPS);


                XpsDocument xpsDocument = new XpsDocument(xpsFilename, FileAccess.Read);
                return xpsDocument;
        }

        private void btnViewDoc_Click(object sender, RoutedEventArgs e)
        {
            string wordDocument = txbSelectedWordFile.Text;
            if (string.IsNullOrEmpty(wordDocument) || !File.Exists(wordDocument))
            {
                MessageBox.Show("The file is invalid. Please select an existing file again.");
            }
            else
            {
                string convertedXpsDoc = string.Concat(System.IO.Path.GetTempPath(), "\\", Guid.NewGuid().ToString(), ".xps");
                XpsDocument xpsDocument = ConvertWordToXps(wordDocument, convertedXpsDoc);
                if (xpsDocument == null)
                {
                    return;
                }

                documentviewWord.Document = xpsDocument.GetFixedDocumentSequence();
            }
        }

    }
}
