namespace BingoCard
{
    public record ExitCode(int Value, string Message)
    {
        public static ExitCode InputFileNotFound(in string filename) => new ExitCode(2, $"Input file {filename} not found.");
        public static ExitCode InputFileTooShort(in string filename, in int numLines) => new ExitCode(3, $"Input file ({filename}) does not contain enough entries. Only found {numLines}.");
    }
}