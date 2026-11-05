using EmlToPst.Models;

namespace EmlToPst.ViewModels;

internal sealed class MailboxSelection(BackupMailbox mailbox) : ObservableObject
{
    private bool selected = true;
    public BackupMailbox Mailbox { get; } = mailbox;
    public string FullName => Mailbox.FullName;
    public bool IsSelected { get => selected; set => Set(ref selected, value); }
}
