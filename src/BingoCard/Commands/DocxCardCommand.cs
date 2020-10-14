using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CliFx;
using CliFx.Attributes;
using LibreOfficeLibrary;
using MailMerge;

namespace BingoCard.Commands
{
    [Command("docx-card")]
    public class DocxCardCommand : ICommand
    {
        [CommandOption("input", 'i', Description = "File containing bingo square contents. One per line.", IsRequired = true)]
        public string InputFilename { get; init; }

        [CommandOption("template", 't', Description = "Docx file containing MailMerge fields to be used as a template for the output bingo card.", IsRequired = true)]
        public string DocxTemplate { get; init; }

        [CommandOption("output", 'o', Description = "Output filename.", IsRequired = false)]
        public string OutputFile { get; init; } = "output.docx";

        [CommandOption("output-format", 'f', Description = "Output format.", IsRequired = false)]
        public OutputFormat OutputFormat { get; init; } = OutputFormat.docx;

        [CommandOption("overwrite-output", 'w', Description = "Overwrite output file if it already exists.", IsRequired = false)]
        public bool OverwriteOutput { get; init; } = false;

        public ValueTask ExecuteAsync(IConsole console)
        {
            if (!File.Exists(InputFilename))
                throw ExitCode.InputFileNotFound(InputFilename).ToCommandException();

            if (!File.Exists(DocxTemplate))
                throw ExitCode.TemplateNotFound(DocxTemplate).ToCommandException();

            if (File.Exists(OutputFile) && !OverwriteOutput)
                throw ExitCode.OutputFileExists(OutputFile).ToCommandException();

            var inputs = File
               .ReadAllLines(InputFilename)
               .Where(l => !string.IsNullOrWhiteSpace(l))
               .TakeRandom(25)
               .ToList();

            if (inputs.Count < 25)
                throw ExitCode.InputFileTooShort(InputFilename, inputs.Count).ToCommandException();

            var card = new CardData(inputs);
            var cells = card.Cells.ToDictionary(c => c.Id.ToLower(), c => c.Value);

            var tmpFile = Path.ChangeExtension(Path.GetTempFileName(), "docx");

            var (mergeSuccess, mergeException) = new MailMerger().Merge(DocxTemplate, cells, tmpFile);

            if (!mergeSuccess)
            {
                throw ExitCode.MailMergeFailed(DocxTemplate, InputFilename, mergeException).ToCommandException();
            }

            if (OutputFormat == OutputFormat.pdf)
            {
                EnsureLibreOffice();

                var tmpPdf = Path.ChangeExtension(Path.GetTempFileName(), "pdf");
                ConvertToPdf(tmpFile, tmpPdf);

                File.Move(tmpPdf, OutputFile, OverwriteOutput);
                File.Delete(tmpPdf);
            }
            else
            {
                File.Move(tmpFile, OutputFile);
            }

            return default;
        }

        private void ConvertToPdf(string tmpDocx, string tmpPdf)
        {
            try
            {
                new DocumentConverter().ConvertToPdf(tmpDocx, tmpPdf);
            }
            catch (Exception e)
            {
                throw ExitCode.PdfConversionFailed(tmpDocx, e).ToCommandException();
            }
        }

        private void EnsureLibreOffice()
        {
            if (!LibreOfficeUtils.IsLibreOfficeOnPath())
            {
                var (found, sofficePath) = LibreOfficeUtils.FindLibreOfficeBinary();

                if (found)
                {
                    var oldValue = Environment.GetEnvironmentVariable("PATH");
                    var newValue = oldValue + Path.PathSeparator + Path.GetDirectoryName(sofficePath);

                    Environment.SetEnvironmentVariable("PATH", newValue);
                }
                else
                {
                    throw ExitCode.LibreOfficeNotFound().ToCommandException();
                }
            }
            if (!LibreOfficeUtils.HasLibreOfficeProfile())
                throw ExitCode.NoLibreOfficeProfile(LibreOfficeUtils.ExpectedProfileDir).ToCommandException();
        }
    }

    public enum OutputFormat
    {
        docx,
        pdf
    }
}