using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CliFx;
using CliFx.Attributes;
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

        public ValueTask ExecuteAsync(IConsole console)
        {
            if (!File.Exists(InputFilename))
                throw ExitCode.InputFileNotFound(InputFilename).ToCommandException();

            if (!File.Exists(DocxTemplate))
                throw ExitCode.TemplateNotFound(DocxTemplate).ToCommandException();

            if (File.Exists(OutputFile))
                throw ExitCode.OutputFileExists(OutputFile).ToCommandException();

            var inputs = File
               .ReadAllLines(InputFilename)
               .Where(l => !string.IsNullOrWhiteSpace(l))
               .TakeRandom(25)
               .ToList();

            if (inputs.Count < 25)
                throw ExitCode.InputFileTooShort(InputFilename, inputs.Count).ToCommandException();

            var card = new CardData(inputs);
            var cells = card.Cells.ToDictionary(c => c.Id, c => c.Value);

            var (mergeSuccess, mergeException) = new MailMerger().Merge(DocxTemplate, cells, OutputFile);

            if (!mergeSuccess)
            {
                throw ExitCode.MailMergeFailed(DocxTemplate, InputFilename, mergeException).ToCommandException();
            }

            return default;
        }
    }
}