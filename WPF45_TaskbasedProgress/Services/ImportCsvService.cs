using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ImportProducts.Models;
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

        public async Task<StringBuilder> GenerateImportCsvAsync(CancellationToken ct, string reff, IEnumerable<string> t2TreFs)
        {
            var tsk = Task.Run(() =>
            {
                var bodyContent = new StringBuilder();
                ct.ThrowIfCancellationRequested();
                var result = _jobs.ProcessT2TRefs(reff, t2TreFs);
                if (result != null)
                {
                    bodyContent.Append(result);
                }
                else
                {
                    return new StringBuilder("");
                }
                    
                return bodyContent;
            }, ct);
            return await tsk;
        }

        public async Task<ObservableCollection<Image>> LoadImagesAsync(CancellationToken ct, IProgress<double> progress, string path)
        {
            var task = Task.Run(() =>
            {
                var fileNames = new ObservableCollection<Image>();
                var images = Directory.GetFiles(path).Select(Path.GetFileName).ToArray();
                foreach (var file in images)
                {
                    fileNames.Add(new Image() { ImageName = file, T2TRef = "" });
                }
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

        public IEnumerable<string> GetuniqueReferenceNumbers(IEnumerable<string> reffs)
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

        public IEnumerable<string> GetT2TRefs(string imagePath)
        {
            return _imageService.ReadImageDetails(imagePath);
        }
    }
}
