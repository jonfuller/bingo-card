using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CliFx;
using CliFx.Attributes;
using LanguageExt;
using static LanguageExt.Prelude;

namespace BingoCard.Commands
{
    [Command("docx-cards")]
    public class DocxCardBatchCommand : ICommand
    {
        [CommandOption("num-cards", 'n', Description = "Number of bingo cards to generate.", IsRequired = true)]
        public int NumCards { get; init; }

        [CommandOption("input", 'i', Description = "File containing bingo square contents. One per line.", IsRequired = true)]
        public string InputFilename { get; init; }

        [CommandOption("template", 't', Description = "Docx file containing MailMerge fields to be used as a template for the output bingo card.", IsRequired = true)]
        public string DocxTemplate { get; init; }

        [CommandOption("output-dir", 'o', Description = "Output directory.", IsRequired = false)]
        public string OutputDirectory { get; init; } = "output";

        [CommandOption("output-format", 'f', Description = "Output format.", IsRequired = false)]
        public OutputFormat OutputFormat { get; init; } = OutputFormat.docx;

        [CommandOption("overwrite-output", 'w', Description = "Overwrite output directory if it already exists.", IsRequired = false)]
        public bool OverwriteOutput { get; init; } = false;

        [CommandOption("collate", 'c', Description = "Collate all cards into a single file named output.pdf. NOTE: only works for PDF output format.", IsRequired = false)]
        public bool Collate { get; init; } = false;

        public ValueTask ExecuteAsync(IConsole console)
        {
            if (!File.Exists(InputFilename))
                throw ExitCode.InputFileNotFound(InputFilename).ToCommandException();

            if (!File.Exists(DocxTemplate))
                throw ExitCode.TemplateNotFound(DocxTemplate).ToCommandException();

            if (Directory.Exists(OutputDirectory) && !OverwriteOutput)
                throw ExitCode.OutputDirectoryExists(OutputDirectory).ToCommandException();
            var inputs = File
               .ReadAllLines(InputFilename)
               .Where(l => !string.IsNullOrWhiteSpace(l))
               .ToList();

            if (inputs.Count < 25)
                throw ExitCode.InputFileTooShort(InputFilename, inputs.Count).ToCommandException();

            var generator = new CardGenerator(inputs);
            var tmpDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());

            Directory.CreateDirectory(tmpDir);

            Action generate = OutputFormat switch
            {
                OutputFormat.docx => () => generator.GenerateDocxs(DocxTemplate, InputFilename, tmpDir, NumCards),
                OutputFormat.pdf => () => generator.GeneratePdfs(DocxTemplate, InputFilename, NumCards, Output()),
                _ => () => { },
            };
            generate();

            if (Directory.Exists(OutputDirectory))
                Directory.Delete(OutputDirectory, true);
            Directory.Move(tmpDir, OutputDirectory);

            return default;

            Either<CardGenerator.Batch, CardGenerator.Collate> Output() => Collate
                ? Right(new CardGenerator.Collate("output.pdf"))
                : Left(new CardGenerator.Batch(tmpDir));
        }
    }
}