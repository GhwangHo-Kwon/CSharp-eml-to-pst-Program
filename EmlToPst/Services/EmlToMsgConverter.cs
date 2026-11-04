using System.IO;
using Aspose.Email.Mapi;

namespace EmlToPst.Services;

internal static class EmlToMsgConverter
{
    public static void Convert(string emlPath, string msgPath)
    {
        using var input = File.OpenRead(emlPath);
        if (input.Length == 0)
        {
            throw new InvalidDataException("비어 있는 EML 파일입니다.");
        }

        using var message = MapiMessage.FromMailMessage(input);
        using var output = new FileStream(msgPath, FileMode.CreateNew, FileAccess.Write);
        message.Save(output);
    }
}
