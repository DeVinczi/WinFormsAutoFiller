using DocumentFormat.OpenXml.Packaging;

using WinFormsAutoFiller.Infrastructure;
using WinFormsAutoFiller.Utilis;

namespace WinFormsAutoFiller.Services
{
    public class DocumentReader
    {
        public async Task<Result<string, Error>> FindWordInWordDocumentAsync(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return new Error("File Not Found", $"The file at path {filePath} does not exist.");
            }

            try
            {
                using (var document = WordprocessingDocument.Open(filePath, false))
                {
                    var body = document.MainDocumentPart?.Document.Body;
                    if (body == null)
                    {
                        return new Error("Document Error", "The document body could not be read.");
                    }

                    var match = RegexPatterns.FindOutHoursRegex().Match(body.InnerText);
                    return match.Success
                        ? match.Groups[1].Value
                        : Errors.WorkHoursAreEmpty;
                }
            }
            catch (Exception ex)
            {
                return new Error("Document Processing Error", ex.Message);
            }
        }
    }
}
