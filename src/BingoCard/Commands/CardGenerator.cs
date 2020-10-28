using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using LanguageExt;
using MailMerge;
using MailMerge.Properties;
using Microsoft.Extensions.Logging.Abstractions;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace BingoCard.Commands
{
    public class CardGenerator
    {
        public record Batch(string OutputDir) { }
        public record Collate(string OutputFile) { }

        private readonly IEnumerable<string> _inputs;
        private readonly MailMerger _mailMerger;

        public CardGenerator(IEnumerable<string> inputs)
        {
            _inputs = inputs;
            _mailMerger = new MailMerger(NullLogger.Instance, new Settings());
        }

        public void GenerateDocx(string DocxTemplate, string InputFilename, string OutputFile, bool OverwriteOutput)
        {
            var tmpFile = DoMailMerge(DocxTemplate, InputFilename);

            File.Move(tmpFile, OutputFile, OverwriteOutput);
        }

        public void GenerateDocxs(string DocxTemplate, string InputFilename, string outputDir, int numCards)
        {
            DoManyMailMerges(numCards, DocxTemplate, InputFilename, outputDir);
        }

        private string GetTempFileName(string extension)
        {
            var filename = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            return Path.ChangeExtension(filename, extension);
        }

        public void GeneratePdf(string DocxTemplate, string InputFilename, string OutputFile, bool OverwriteOutput)
        {
            var tmpFile = DoMailMerge(DocxTemplate, InputFilename);

            EnsureLibreOffice();

            var tmpPdf = GetTempFileName("pdf");
            ConvertToPdf(tmpFile, tmpPdf);
            File.Move(tmpPdf, OutputFile, OverwriteOutput);
            File.Delete(tmpFile);
        }

        public void GeneratePdfs(string DocxTemplate, string InputFilename, int numCards, Either<Batch, Collate> output)
        {
            EnsureLibreOffice();
            
            output.Match(
                Left: l =>
                {
                    var docs = DoManyMailMerges(numCards, DocxTemplate, InputFilename, l.OutputDir).ToList();
                    ConvertToPdfs(docs, l.OutputDir);
                    docs.ToList().ForEach(File.Delete);
                },
                Right: r =>
                {
                    var outputDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
                    Directory.CreateDirectory(outputDir);
                    var docs = DoManyMailMerges(numCards, DocxTemplate, InputFilename, outputDir).ToList();
                    ConvertToPdfs(docs, outputDir);
                    MergePdfs(Directory.GetFiles(outputDir, "*.pdf"), r.OutputFile);
                    Directory.Delete(outputDir, recursive: true);
                });
        }

        private void MergePdfs(IEnumerable<string> pdfs, string outputFile)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            using (var output = new PdfDocument())
            {
                foreach (var pdfFilename in pdfs)
                {
                    using var input = PdfReader.Open(pdfFilename, PdfDocumentOpenMode.Import);
                    CopyPages(input, output);
                }
                output.Save(outputFile);
            }

            void CopyPages(PdfDocument from, PdfDocument to)
            {
                foreach (var page in from.Pages)
                {
                    to.AddPage(page);
                }
            }
        }

        private IEnumerable<string> DoManyMailMerges(int count, string DocxTemplate, string InputFilename, string outputDir)
        {
            for (var i=0; i<count; i++)
            {
                var card = new CardData(_inputs.TakeRandom(25));
                var cells = card.Cells.ToDictionary(c => c.Id.ToLower(), c => c.Value);
                var filename = Path.Combine(outputDir, $"{i}.docx");
                var (mergeSuccess, mergeException) = _mailMerger.Merge(DocxTemplate, cells, filename);

                if (!mergeSuccess)
                {
                    throw ExitCode.MailMergeFailed(DocxTemplate, InputFilename, mergeException).ToCommandException();
                }

                yield return filename;
            }
        }

        private string DoMailMerge(string DocxTemplate, string InputFilename)
        {
            var card = new CardData(_inputs.TakeRandom(25));
            var cells = card.Cells.ToDictionary(c => c.Id.ToLower(), c => c.Value);

            var tmpFile = GetTempFileName("docx");

            var (mergeSuccess, mergeException) = _mailMerger.Merge(DocxTemplate, cells, tmpFile);

            if (!mergeSuccess)
            {
                throw ExitCode.MailMergeFailed(DocxTemplate, InputFilename, mergeException).ToCommandException();
            }

            return tmpFile;
        }

        private void ConvertToPdfs(IEnumerable<string> docs, string outputDir, string profile = null)
        {
            try
            {
                LibreOfficeUtils.ConvertFiles(docs, outputDir, profile);
            }
            catch (Exception e)
            {
                throw ExitCode.PdfBatchConversionFailed(outputDir, e).ToCommandException();
            }
        }

        private void ConvertToPdf(string tmpFile, string tmpDocx, string profile = null)
        {
            try
            {
                LibreOfficeUtils.Convert(tmpDocx, tmpFile, profile);
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
}