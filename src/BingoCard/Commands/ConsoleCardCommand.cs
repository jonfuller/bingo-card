using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CliFx;
using CliFx.Attributes;
using ConsoleTables;

namespace BingoCard.Commands
{
    [Command("console-card")]
    public class ConsoleCardCommand : ICommand
    {
        [CommandOption("input", 'i', Description = "File containing bingo square contents. One per line.", IsRequired = true)]
        public string InputFilename { get; init; }

        public ValueTask ExecuteAsync(IConsole console)
        {
            if (!File.Exists(InputFilename))
                throw ExitCode.InputFileNotFound(InputFilename).ToCommandException();

            var inputs = File
               .ReadAllLines(InputFilename)
               .Where(l => !string.IsNullOrWhiteSpace(l))
               .TakeRandom(25)
               .ToList();

            if (inputs.Count < 25)
                throw ExitCode.InputFileTooShort(InputFilename, inputs.Count).ToCommandException();

            var card = new CardData(inputs);

            ConsoleTable
               .From(card.Rows)
               .Configure(o =>
                {
                    o.OutputTo = console.Output;
                    o.EnableCount = false;
                })
               .Write();

            return default;
        }
    }
}