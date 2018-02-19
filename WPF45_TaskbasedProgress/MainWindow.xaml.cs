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
using ImportProducts.Models;
using ImportProducts.Services;

namespace WPF45_TaskbasedProgress
{
    /// <inheritdoc cref="" />
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        CancellationTokenSource _cancelToken;
        IProgress<double> _progressOperation;
        private readonly ImportCsvService _importCsvService;
        public static string ImagePath;
        private StringBuilder _csv;

        public MainWindow()
        {
            InitializeComponent();
            btnCancel.IsEnabled = false;
            _csv = new StringBuilder();
            _importCsvService = new ImportCsvService();
        }

        // Displaying Employees in DataGrid 
        private async void btnLoadImages_Click(object sender, RoutedEventArgs e)
        {
            _cancelToken = new CancellationTokenSource();
            btnLoadImages.IsEnabled = false;
            btnCancel.IsEnabled = true;
            TxtStatus.Text = "Loading.....";
            _progressOperation = new Progress<double>(value => progress.Value = value);

            try
            {
                var folderBrowserDlg = new FolderBrowserDialog
                {
                    ShowNewFolderButton = true
                };
                var dlgResult = folderBrowserDlg.ShowDialog();

                ImagePath = folderBrowserDlg.SelectedPath;

                var emps = await _importCsvService.LoadImagesAsync(_cancelToken.Token, _progressOperation, ImagePath);

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
                _cancelToken.Dispose();
                btnLoadImages.IsEnabled = true;
                btnCancel.IsEnabled = false;
            }
        }

        private async void btnGenerateImportCsv_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _csv = new StringBuilder();
                _cancelToken = new CancellationTokenSource();
                btnLoadImages.IsEnabled = false;
                btnGenerateImportCsv.IsEnabled = false;
                btnCancel.IsEnabled = true;
                TxtStatus.Text = "Generating.....";
                _progressOperation = (IProgress<double>)new Progress<double>(value => progress.Value = value);

                TxtStatus.Text = "Removing exisiting import product file...";
                await _importCsvService.CleanupExistingFile();
                TxtStatus.Text = "Generating import product csv file, please wait this can take several minutes....";

                var t2TreFs = _importCsvService.GetT2TRefs(ImagePath);
                var uniqueReffs = _importCsvService.GetuniqueReferenceNumbers(t2TreFs);
                var recCount = 0;

                foreach (var reff in uniqueReffs)
                {
                    var result = await _importCsvService.GenerateImportCsvAsync(_cancelToken.Token, reff, t2TreFs);

                    if(string.IsNullOrEmpty(result.ToString()))
                        dgError.Items.Add(new Error(){RefNumber = reff.Substring(0,9)});
                    else
                        _csv.AppendLine(result.ToString());

                    ++recCount;
                    _progressOperation.Report(recCount * 100.0 / 70);
                }

                _progressOperation.Report(100);
                File.AppendAllText(System.Configuration.ConfigurationManager.AppSettings["OutputPath"], _csv.ToString());
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
                _cancelToken.Dispose();
                btnGenerateImportCsv.IsEnabled = true;
                btnCancel.IsEnabled = false;
                btnLoadImages.IsEnabled = true;
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            _cancelToken.Cancel();
        }
    }
}
