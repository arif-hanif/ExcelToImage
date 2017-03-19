using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Spire.Xls;

namespace ExcelToImage
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string fileXls;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog for Excel spreadsheet
            Microsoft.Win32.OpenFileDialog dlgXls = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlgXls.DefaultExt = ".xlsx";
            dlgXls.Filter = "XLSX Files (*.xlsx)|*.xlsx|XLS Files (*.xls)|*.xls";

            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlgXls.ShowDialog();

            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                fileXls = dlgXls.FileName;
            }

            // Get path of Excel file
            string fileXlsFolder = fileXls.Substring(0, fileXls.LastIndexOf((@"\")) + 1);

            // Get Excel filename to use for image filename
            int fileXlsLength = fileXls.Length;
            int fileXlsExtensionLength = fileXlsLength - (fileXls.Substring(0, fileXls.LastIndexOf((@"."))).Length + 1);
            int fileXlsEndOfFilename = fileXls.Substring(0, fileXls.LastIndexOf((@"."))).Length;
            int fileXlsFilenameLength = (fileXlsFolder.Length + 1) - (fileXls.Length - fileXlsEndOfFilename);
            string fileXlsFilename = fileXls.Substring(fileXlsFolder.Length, fileXlsEndOfFilename - fileXlsFolder.Length);

            // Create Workbook instance and load file
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(fileXls);
            Worksheet sheet = workbook.Worksheets[0];

            //Hiding the gridlines for the first worksheet
            sheet.GridLinesVisible = false;

            // Save to image
            string savedImage = fileXlsFolder + fileXlsFilename + ".png";
            sheet.SaveToImage(savedImage, 1, 1, 7, 14);

            //// TEMP: Dialog for displaying results
            //MessageBoxResult tempResult = System.Windows.MessageBox.Show(
            //    "fileXlsLength: " + fileXlsLength.ToString() + "\n" +
            //    "fileXlsExtensionLength: " + fileXlsExtensionLength.ToString() + "\n" +
            //    "fileXlsEndOfFilename: " + fileXlsEndOfFilename.ToString() + "\n" +
            //    "fileXlsFilename: " + fileXlsFilename + "\n" +
            //    "fileXlsFilenameLength: " + fileXlsFilenameLength + "\n" +
            //    "fileXlsFolder.Length: " + fileXlsFolder.Length + "\n" +
            //    "Length of fileXlsFolder: " + fileXlsFolder.Length.ToString()
            // );

            textBox.Text = "Schedule exported as image: " + savedImage;

            //// Open image
            //System.Diagnostics.Process.Start(@"C:\Users\J. Merlan\Dev Lab\_Test Content\AHU Schedule.bmp");


        }
    }
}
