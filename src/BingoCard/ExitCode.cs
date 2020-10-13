using System;

namespace BingoCard
{
    public record ExitCode(int Value, string Message)
    {
        public static ExitCode InputFileNotFound(in string filename) =>
            new ExitCode(2, $"Input file {filename} not found.");
        public static ExitCode InputFileTooShort(in string filename, in int numLines) =>
            new ExitCode(3, $"Input file ({filename}) does not contain enough entries. Only found {numLines}.");
        public static ExitCode TemplateNotFound(string docxTemplate) =>
            new ExitCode(4, $"Docx Template file {docxTemplate} not found.");
        public static ExitCode MailMergeFailed(string docxTemplate, string inputFilename, AggregateException exception) =>
            new ExitCode(5, $"MailMerge failed for template file {docxTemplate} and input file {inputFilename}. Exception details:{Environment.NewLine}{exception}");
        public static ExitCode OutputFileExists(string outputFile) =>
            new ExitCode(6, $"Output file {outputFile} already exists. Output file will not be overwritten.");
    }
}