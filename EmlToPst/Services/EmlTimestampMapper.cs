using System.Globalization;
using System.Text;
using OfficeIMO.Email;

namespace EmlToPst.Services;

internal static class EmlTimestampMapper
{
    private static readonly string[] DateFormats = ["d MMM yyyy H:mm:ss zzz", "d MMM yyyy H:mm zzz"];

    public static string Apply(EmailDocument document)
    {
        if (!document.Date.HasValue)
        {
            EmailHeader? sentHeader = document.Headers.FirstOrDefault(
                header => header.Name.Equals("Date", StringComparison.OrdinalIgnoreCase));
            if (sentHeader is not null && TryParse(RemoveComments(sentHeader.RawValue ?? sentHeader.Value), out var sent))
            {
                document.Date = sent;
            }
        }

        int hop = 0;
        foreach (EmailHeader header in document.Headers)
        {
            if (!header.Name.Equals("Received", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            hop++;
            string trace = RemoveComments(header.RawValue ?? header.Value);
            int separator = trace.LastIndexOf(';');
            if (separator >= 0 && TryParse(trace[(separator + 1)..], out var received))
            {
                document.ReceivedDate = received;
                return $"Received[{hop}]";
            }

            if (hop == 1 && document.ReceivedDate.HasValue)
            {
                return "library";
            }
        }

        return document.ReceivedDate.HasValue ? "library" : "none";
    }

    private static bool TryParse(string value, out DateTimeOffset timestamp)
    {
        timestamp = default;
        string date = value.Trim();
        int zoneStart = date.LastIndexOfAny([' ', '\t', '\r', '\n']);
        if (zoneStart < 0)
        {
            return false;
        }

        string zone = date[(zoneStart + 1)..].ToUpperInvariant();
        zone = zone switch
        {
            "UT" or "UTC" or "GMT" or "Z" => "+00:00",
            "EST" or "CDT" => "-05:00",
            "EDT" => "-04:00",
            "CST" or "MDT" => "-06:00",
            "MST" or "PDT" => "-07:00",
            "PST" => "-08:00",
            _ when zone.Length == 5 && zone[0] is '+' or '-' => zone.Insert(3, ":"),
            _ => zone
        };

        string dayAndTime = date[..zoneStart].TrimEnd();
        int weekdaySeparator = dayAndTime.IndexOf(',');
        if (weekdaySeparator >= 0)
        {
            dayAndTime = dayAndTime[(weekdaySeparator + 1)..].TrimStart();
        }

        return DateTimeOffset.TryParseExact($"{dayAndTime} {zone}", DateFormats, CultureInfo.InvariantCulture,
            DateTimeStyles.AllowWhiteSpaces, out timestamp);
    }

    private static string RemoveComments(string header)
    {
        var value = new StringBuilder(header.Length);
        int depth = 0;
        bool quoted = false;

        for (int index = 0; index < header.Length; index++)
        {
            char character = header[index];
            if (character == '\\' && (quoted || depth > 0) && index + 1 < header.Length)
            {
                index++;
                if (depth == 0)
                {
                    value.Append(character).Append(header[index]);
                }

                continue;
            }

            if (depth > 0)
            {
                if (character == '(')
                {
                    depth++;
                }
                else if (character == ')')
                {
                    depth--;
                }

                continue;
            }

            if (character == '(' && !quoted)
            {
                depth = 1;
                value.Append(' ');
                continue;
            }

            if (character == '"')
            {
                quoted = !quoted;
            }

            value.Append(character);
        }

        return value.ToString();
    }
}
