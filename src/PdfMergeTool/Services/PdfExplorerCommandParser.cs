namespace PdfMergeTool.Services;

public static class PdfExplorerCommandParser
{
    public static PdfExplorerCommand Parse(IEnumerable<string> arguments)
    {
        foreach (var argument in arguments)
        {
            if (string.Equals(argument, "--split-pages", StringComparison.OrdinalIgnoreCase))
            {
                return PdfExplorerCommand.PageSplit;
            }

            if (string.Equals(argument, "--split-interval", StringComparison.OrdinalIgnoreCase))
            {
                return PdfExplorerCommand.IntervalSplit;
            }

            if (string.Equals(argument, "--split-parity", StringComparison.OrdinalIgnoreCase))
            {
                return PdfExplorerCommand.ParitySplit;
            }
        }

        return PdfExplorerCommand.None;
    }
}

public enum PdfExplorerCommand
{
    None,
    PageSplit,
    IntervalSplit,
    ParitySplit
}
