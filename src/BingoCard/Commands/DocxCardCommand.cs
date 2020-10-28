using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CliFx;
using CliFx.Attributes;

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
               .ToList();

            if (inputs.Count < 25)
                throw ExitCode.InputFileTooShort(InputFilename, inputs.Count).ToCommandException();

            var generator = new CardGenerator(inputs);

            Action generate = OutputFormat switch
            {
                OutputFormat.docx => () => generator.GenerateDocx(DocxTemplate, InputFilename, OutputFile, OverwriteOutput),
                OutputFormat.pdf => () => generator.GeneratePdf(DocxTemplate, InputFilename, OutputFile, OverwriteOutput),
                _ => () => { },
            };
            generate();

            return default;
        }
    }

    public enum OutputFormat
    {
        docx,
        pdf
    }
}