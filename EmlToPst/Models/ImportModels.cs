namespace EmlToPst.Models;

internal sealed record ImportOptions(
    string SourceFolder,
    string OutputPath,
    bool IncludeSubfolders,
    bool AttachToOutlook);

internal sealed record ImportProgress(
    int Processed,
    int Total,
    int Imported,
    int Failed,
    string CurrentFile);

internal sealed record ImportFailure(string FilePath, string Reason);

internal sealed record ImportResult(
    int Total,
    int Imported,
    IReadOnlyList<ImportFailure> Failures,
    bool Cancelled,
    string? PstPath,
    bool AttachedToOutlook,
    IReadOnlyList<string> Warnings);
