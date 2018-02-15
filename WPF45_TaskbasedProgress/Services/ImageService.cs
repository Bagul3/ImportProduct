using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ImportProducts.Services
{
    public class ImageService
    {
        private readonly LogWriter _logger;

        public ImageService()
        {
           _logger  = new LogWriter();
        }

        public IEnumerable<string> ReadImageDetails(string path)
        {
            try
            {
                var imageDetails = Directory.GetFiles(path, "*.jpg*", SearchOption.AllDirectories)
                    .ToList();
                return ParseImageNames(imageDetails);
            }
            catch (Exception e)
            {
                _logger.LogWrite("Error occured reading files: " + e);
                return new List<string>();
            }
        }

        private static IEnumerable<string> ParseImageNames(IEnumerable<string> imageDetails)
        {
            return imageDetails.Select(Path.GetFileNameWithoutExtension).ToList();
        }
    }
}
