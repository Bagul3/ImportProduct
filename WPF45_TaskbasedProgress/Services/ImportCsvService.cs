using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WPF45_TaskbasedProgress;

namespace ImportProducts.Services
{
    public class ImportCsvService
    {
        private readonly CsvJobs _jobs;
        private readonly ImageService _imageService;

        public ImportCsvService()
        {
            this._jobs = new CsvJobs();
            this._imageService = new ImageService();
        }

        public StringBuilder GenerateImportCsv(CancellationToken ct, IProgress<double> progress, string imagePath)
        {
            var bodyContent = new StringBuilder();
            var recCount = 0;
            var t2TreFs = GetT2TRefs(imagePath);
            var uniqueReffs = GetuniqueReferenceNumbers(t2TreFs);

            Parallel.ForEach(uniqueReffs, (reff) =>
            {
                ct.ThrowIfCancellationRequested();
                bodyContent.Append(_jobs.ProcessT2TRefs(reff, t2TreFs));
                ++recCount;
                progress.Report(recCount * 100.0 / 50);
            });

            return bodyContent;
        }

        public async Task<ObservableCollection<Image>> LoadImagesAsync(CancellationToken ct, IProgress<double> progress, string path)
        {
            var task = Task.Run(() => {
                var fileNames = new ObservableCollection<Image>();
                var images = Directory.GetFiles(path).Select(Path.GetFileName).ToArray();
                Parallel.ForEach(images, (file) =>
                {
                    fileNames.Add(new Image() {ImageName = file, T2TRef = ""});
                });

                return fileNames;
            }, ct);

            return await task;
        }

        public async Task CleanupExistingFile()
        {
            var tsk = Task.Run(() =>
            {
                _jobs.DoCleanup();
            });

            await tsk;
        }

        private IEnumerable<string> GetuniqueReferenceNumbers(IEnumerable<string> reffs)
        {
            var checkNumber = "0000";
            var uniqueReferenceNumbers = new List<string>();
            foreach (var reff in reffs)
            {
                if (!reff.Contains(checkNumber))
                {
                    uniqueReferenceNumbers.Add(reff);
                    checkNumber = reff.Substring(0, 9);
                }
            }
            return uniqueReferenceNumbers;
        }

        private IEnumerable<string> GetT2TRefs(string imagePath)
        {
            return _imageService.ReadImageDetails(imagePath);
        }
    }
}
