using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using ImportProducts.Services;

namespace WPF45_TaskbasedProgress
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        CancellationTokenSource cancelToken;
        IProgress<double> progressOperation;
        private readonly ImportCsvService _importCsvService;
        public static string ImagePath;
        private StringBuilder csv;

        public MainWindow()
        {
            InitializeComponent();
            btnCancel.IsEnabled = false;
            csv = new StringBuilder();
            _importCsvService = new ImportCsvService();
        }

        // Displaying Employees in DataGrid 
        private async void btnLoadImages_Click(object sender, RoutedEventArgs e)
        {
            cancelToken = new CancellationTokenSource();
            btnLoadImages.IsEnabled = false;
            btnCancel.IsEnabled = true;
            TxtStatus.Text = "Loading.....";
            progressOperation = new Progress<double>(value => progress.Value = value);

            try
            {
                var folderBrowserDlg = new FolderBrowserDialog
                {
                    ShowNewFolderButton = true
                };
                var dlgResult = folderBrowserDlg.ShowDialog();

                ImagePath = folderBrowserDlg.SelectedPath;
                
                var emps = await _importCsvService.LoadImagesAsync(cancelToken.Token, progressOperation, ImagePath);

                foreach (var item in emps)
                {
                    dgEmp.Items.Add(item);
                }

                TxtStatus.Text = "Operation Completed";
            }
            catch (OperationCanceledException ex)
            {
                TxtStatus.Text = "Operation cancelled" + ex.Message;
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "Operation cancelled" + ex.Message;
            }
            finally
            {
                cancelToken.Dispose();
                btnLoadImages.IsEnabled = true;
                btnCancel.IsEnabled = false;
            }
        }

        private async void btnGenerateImportCsv_Click(object sender, RoutedEventArgs e)
        {
            csv = new StringBuilder();
            cancelToken = new CancellationTokenSource();
            btnLoadImages.IsEnabled = false;
            btnGenerateImportCsv.IsEnabled = false;
            btnCancel.IsEnabled = true;
            TxtStatus.Text = "Generating.....";
            progressOperation = (IProgress<double>)new Progress<double>(value => progress.Value = value);
            
            try
            {
                TxtStatus.Text = "Removing exisiting import product file...";
                await _importCsvService.CleanupExistingFile();
                TxtStatus.Text = "Generating import product csv file, please wait this can take several minutes....";
                
                var result = _importCsvService.GenerateImportCsv(cancelToken.Token, progressOperation, ImagePath);
                csv.AppendLine(result.ToString());
                progressOperation.Report(100);
                File.AppendAllText(System.Configuration.ConfigurationManager.AppSettings["OutputPath"], csv.ToString());
                TxtStatus.Text = "Operation completed";
            }
            catch (OperationCanceledException ex)
            {
                TxtStatus.Text = "Operation cancelled" + ex.Message;
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "Operation cancelled" + ex.Message;
            }
            finally
            {
                cancelToken.Dispose();
                btnGenerateImportCsv.IsEnabled = true;
                btnCancel.IsEnabled = false;
                btnLoadImages.IsEnabled = true;
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            cancelToken.Cancel();
        }
    }
}
