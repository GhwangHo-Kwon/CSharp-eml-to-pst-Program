using System.IO;
using System.Net.Sockets;
using EmlToPst.Models;
using MailKit;
using MailKit.Net.Imap;

namespace EmlToPst.Services;

internal sealed class ImapConnectionRunner(ImapConnectionOptions options, string password, ImportLog log,
    Func<IImapBackupSource> factory, Func<TimeSpan, CancellationToken, Task>? delay = null) : IAsyncDisposable
{
    private static readonly int[] RetrySeconds = [2, 5, 10];
    private readonly Func<TimeSpan, CancellationToken, Task> pause = delay ?? ((duration, cancellation) => Task.Delay(duration, cancellation));
    private IImapBackupSource? source;

    public async Task<T> RunAsync<T>(string? mailbox, uint? expectedValidity,
        Func<IImapBackupSource, uint, Task<T>> operation, CancellationToken token)
    {
        for (int attempt = 0; ; attempt++)
        {
            token.ThrowIfCancellationRequested();
            try
            {
                if (source is null)
                {
                    source = factory();
                    log.Info("IMAP_CONNECT", $"host={options.Host} | port={options.Port} | account={options.Account} | tls=SslOnConnect");
                    await source.ConnectAsync(options, password, token).ConfigureAwait(false);
                    log.Info("IMAP_CONNECTED", "authenticated=true");
                }
                uint validity = mailbox is null ? 0 : await source.OpenMailboxAsync(mailbox, token).ConfigureAwait(false);
                if (expectedValidity.HasValue && expectedValidity.Value != validity)
                    throw new UidValidityChangedException(mailbox!);
                return await operation(source, validity).ConfigureAwait(false);
            }
            catch (Exception ex) when (!token.IsCancellationRequested && IsTransient(ex) && attempt < RetrySeconds.Length)
            {
                log.Error("IMAP_RETRY", $"retry={attempt + 1}/{RetrySeconds.Length} | mailbox={mailbox} | waitSec={RetrySeconds[attempt]}", ex);
                await ResetAsync().ConfigureAwait(false);
                await pause(TimeSpan.FromSeconds(RetrySeconds[attempt]), token).ConfigureAwait(false);
            }
            catch (Exception ex) when (!token.IsCancellationRequested && IsTransient(ex))
            {
                await ResetAsync().ConfigureAwait(false);
                token.ThrowIfCancellationRequested();
                // Do not spend another full retry budget on every remaining message during an outage.
                throw new ImapTransferFailedException(ex);
            }
        }
    }

    private static bool IsTransient(Exception ex) => ex is IOException or SocketException or TimeoutException
        or ServiceNotConnectedException or ImapProtocolException;

    private async Task ResetAsync()
    {
        IImapBackupSource? old = source;
        source = null;
        if (old is null) return;
        try { await old.DisposeAsync().ConfigureAwait(false); }
        catch (Exception ex) { log.Error("IMAP_DISCONNECT", "disconnect during cleanup failed", ex); }
        log.Info("IMAP_DISCONNECT", "connection released");
    }

    public ValueTask DisposeAsync() => new(ResetAsync());
}
