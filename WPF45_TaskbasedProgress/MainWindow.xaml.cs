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
using ImportProducts.Services;

namespace WPF45_TaskbasedProgress
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ObservableCollection<Employee> Employees;
        ObservableCollection<Image> fileNames;
        CancellationTokenSource cancelToken;
        IProgress<double> progressOperation;
        public static string imagePath;
        private ImportCsvJob job;
        private StringBuilder csv;

        public MainWindow()
        {
            InitializeComponent();
            fileNames = new ObservableCollection<Image>();
            Employees = new ObservableCollection<Employee>();
            btnCancel.IsEnabled = false;
            job = new ImportCsvJob();
            csv = new StringBuilder();
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

                imagePath = folderBrowserDlg.SelectedPath;
                
                var Emps = await LoadImagesAsync(cancelToken.Token, progressOperation);

                foreach (var item in Emps)
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


        // Async Method to Load Employees

        async Task<ObservableCollection<Image>> LoadImagesAsync(CancellationToken ct, IProgress<double> progress)
        {
            Employees.Clear();
            var task = Task.Run(() => {
                var images = Directory.GetFiles(imagePath).Select(Path.GetFileName).ToArray();
                foreach (var file in images)
                {
                    fileNames.Add(new Image() {ImageName = file , T2TRef = ""});
                }

                return fileNames;
            });

            return await task;
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
            var headers = $"{"store"},{"websites"},{"attribut_set"},{"type"},{"sku"},{"has_options"},{"name"},{"page_layout"},{"options_container"},{"price"},{"weight"},{"status"},{"visibility"},{"short_description"},{"qty"},{"product_name"},{"color"}," +
                          $"{"size"},{"tax_class_id"},{"configurable_attributes"},{"simples_skus"},{"manufacturer"},{"is_in_stock"},{"categories"},{"season"},{"stock_type"},{"image"},{"small_image"},{"thumbnail"},{"gallery"}," +
                          $"{"condition"},{"ean"},{"description"},{"model"}";

            csv.AppendLine(headers);
            try
            {
                TxtStatus.Text = "Removing exisiting import product file...";
                await CleanupExistingFile();
                TxtStatus.Text = "Generating import product csv file, please wait this can take several minutes....";
                var t2tRefs = new ImageService().ReadImageDetails(imagePath);
                var result = await GenerateImportCsv(t2tRefs, cancelToken.Token, progressOperation);
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
                        bodyContent.Append(job.DoJob(refff, t2TreFs));
                        checkNumber = refff.Substring(0,9);
                        ++recCount;
                        progress.Report(recCount * 100.0 / 50);
                    }
                }
                return bodyContent;
            });

            return await tsk;
        }
        

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            cancelToken.Cancel();
        }
    }
}
