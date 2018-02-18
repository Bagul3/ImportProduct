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
using System.Xml.Linq;
using ImportProducts.Models;
using ImportProducts.Services;

namespace WPF45_TaskbasedProgress
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        readonly ObservableCollection<Image> _fileNames;
        private ObservableCollection<Error> _errors;
        CancellationTokenSource _cancelToken;
        IProgress<double> _progressOperation;
        public static string ImagePath;
        private readonly ImportCsvJob job;
        private StringBuilder _csv;

        public MainWindow()
        {
            InitializeComponent();
            _fileNames = new ObservableCollection<Image>();
            btnCancel.IsEnabled = false;
            job = new ImportCsvJob();
            _csv = new StringBuilder();
            _errors = new ObservableCollection<Error>();
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
                
                var Emps = await LoadImagesAsync(_cancelToken.Token, _progressOperation);

                foreach (var item in Emps)
                {
                    dgEmp.Items.Add(item);
                }
                _errors.Clear();
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


        // Async Method to Load Employees

        private async Task<ObservableCollection<Image>> LoadImagesAsync(CancellationToken ct, IProgress<double> progress)
        {
            var task = Task.Run(() => {
                var images = Directory.GetFiles(ImagePath).Select(Path.GetFileName).ToArray();
                foreach (var file in images)
                {
                    _fileNames.Add(new Image() {ImageName = file , T2TRef = ""});
                }

                return _fileNames;
            });

            return await task;
        }

        private async void btnGenerateImportCsv_Click(object sender, RoutedEventArgs e)
        {
            _csv = new StringBuilder();
            _cancelToken = new CancellationTokenSource();
            btnLoadImages.IsEnabled = false;
            btnGenerateImportCsv.IsEnabled = false;
            btnCancel.IsEnabled = true;
            TxtStatus.Text = "Generating.....";
            _progressOperation = (IProgress<double>)new Progress<double>(value => progress.Value = value);
            var headers = $"{"store"},{"websites"},{"attribut_set"},{"type"},{"sku"},{"has_options"},{"name"},{"page_layout"},{"options_container"},{"price"},{"weight"},{"status"},{"visibility"},{"short_description"},{"qty"},{"product_name"},{"color"}," +
                          $"{"size"},{"tax_class_id"},{"configurable_attributes"},{"simples_skus"},{"manufacturer"},{"is_in_stock"},{"categories"},{"season"},{"stock_type"},{"image"},{"small_image"},{"thumbnail"},{"gallery"}," +
                          $"{"condition"},{"ean"},{"description"},{"model"}";

            _csv.AppendLine(headers);
            try
            {
                TxtStatus.Text = "Removing exisiting import product file...";
                await CleanupExistingFile();
                TxtStatus.Text = "Generating import product csv file, please wait this can take several minutes....";
                var t2tRefs = new ImageService().ReadImageDetails(ImagePath);
                var result = await GenerateImportCsv(t2tRefs, _cancelToken.Token, _progressOperation);
                _csv.AppendLine(result.ToString());
                _progressOperation.Report(100);
                foreach (var error in _errors)
                {
                    dgError.Items.Add(error);
                }
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



        // Calcluate the Tax for every Employee Record.
        private async Task CleanupExistingFile()
        {
            var tsk = Task.Run(() =>
            {
                job.DoCleanup();
            });

            await tsk;
        }

        private async Task<StringBuilder> GenerateImportCsv(IEnumerable<string> t2TreFs, CancellationToken ct, IProgress<double> progress)
        {
            var bodyContent = new StringBuilder();
            var checkNumber = "00000";
            int recCount = 0;
            var tsk = Task.Run(() =>
            {
                foreach (var refff in t2TreFs.Where(x => !x.Contains("_")))
                {
                    ct.ThrowIfCancellationRequested();

                    if (!refff.Contains(checkNumber))
                    {
                        bodyContent.Append(job.DoJob(refff, t2TreFs, ref _errors));
                        checkNumber = refff.Substring(0,9);
                        ++recCount;
                        progress.Report(recCount * 100.0 / 80);
                    }
                }
                return bodyContent;
            });

            return await tsk;
        }
        

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            _cancelToken.Cancel();
        }
    }
}
