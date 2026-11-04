using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace EmlToPst.Services;

internal static class OutlookPstAttacher
{
    private const int OlStoreUnicode = 2;

    public static void Attach(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("연결할 PST를 찾을 수 없습니다.", path);
        }

        if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
        {
            throw new InvalidOperationException("Outlook 연결은 STA 스레드에서 실행해야 합니다.");
        }

        Type outlookType = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException("클래식 Outlook을 설치하고 기본 프로필을 설정해 주세요.");
        object application = Activator.CreateInstance(outlookType)
            ?? throw new InvalidOperationException("클래식 Outlook을 시작할 수 없습니다.");
        object? session = null;
        try
        {
            session = Call(application, "GetNamespace", "MAPI")
                ?? throw new InvalidOperationException("Outlook 프로필을 열 수 없습니다.");
            Call(session, "AddStoreEx", path, OlStoreUnicode);
        }
        finally
        {
            if (session is not null)
            {
                Marshal.ReleaseComObject(session);
            }

            Marshal.ReleaseComObject(application);
        }
    }

    private static object? Call(object target, string name, params object[] arguments)
    {
        return target.GetType().InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.DoNotWrapExceptions,
            null, target, arguments, CultureInfo.InvariantCulture);
    }
}
