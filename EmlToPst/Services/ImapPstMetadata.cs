using EmlToPst.Models;
using MailKit;
using OfficeIMO.Email;

namespace EmlToPst.Services;

internal static class ImapPstMetadata
{
    public static (string ReceivedSource, bool? IsRead) Apply(EmailDocument document, StoredBackupMessage? metadata)
    {
        string source = EmlTimestampMapper.Apply(document);
        if (metadata is null) return (source, null);
        if (metadata.InternalDate.HasValue)
        {
            document.ReceivedDate = metadata.InternalDate;
            source = "IMAP.InternalDate";
        }
        document.Date ??= metadata.SentDate;
        bool isRead = (metadata.Flags & MessageFlags.Seen) != 0;
        document.MessageMetadata.IsRead = isRead;
        document.MessageMetadata.IsDraft = (metadata.Flags & MessageFlags.Draft) != 0;
        return (source, isRead);
    }
}
