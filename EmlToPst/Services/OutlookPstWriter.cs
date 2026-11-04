using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace EmlToPst.Services;

internal sealed class OutlookPstWriter : IDisposable
{
    private const int OlStoreUnicode = 2;
    private const int OlFolderInbox = 6;
    private const int OlMail = 43;
    private const int OlDiscard = 1;

    private readonly object application;
    private readonly object session;
    private object? storeRoot;
    private object? importRoot;
    private object? destination;
    private string? destinationPath;
    private string? workingPath;

    private OutlookPstWriter(object application, object session)
    {
        this.application = application;
        this.session = session;
    }

    public static OutlookPstWriter Connect()
    {
        if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
        {
            throw new InvalidOperationException("Outlook 호출은 WPF STA 스레드에서 실행해야 합니다.");
        }

        Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
        if (outlookType is null)
        {
            throw new InvalidOperationException("클래식 Outlook이 설치된 Windows PC에서 실행해 주세요.");
        }

        object application = Activator.CreateInstance(outlookType)
            ?? throw new InvalidOperationException("클래식 Outlook을 시작할 수 없습니다.");
        try
        {
            object session = Call(application, "GetNamespace", "MAPI")
                ?? throw new InvalidOperationException("Outlook 프로필을 열 수 없습니다.");
            return new OutlookPstWriter(application, session);
        }
        catch
        {
            Release(application);
            throw;
        }
    }

    public void Create(string path)
    {
        // The staging path belongs to this run, so only this store may be detached.
        workingPath = path;
        Call(session, "AddStoreEx", path, OlStoreUnicode);
        storeRoot = FindStoreRoot(path)
            ?? throw new InvalidOperationException("Outlook에서 생성한 PST를 찾을 수 없습니다.");

        importRoot = OpenOrCreateChild(storeRoot, "EML 가져오기");
    }

    public void Import(string msgPath, string relativeFolder)
    {
        if (importRoot is null)
        {
            throw new InvalidOperationException("PST의 가져오기 폴더를 열지 못했습니다.");
        }

        if (!string.Equals(destinationPath, relativeFolder, StringComparison.OrdinalIgnoreCase))
        {
            Release(destination);
            destination = null;
            destinationPath = null;

            if (relativeFolder != ".")
            {
                destination = OpenRelativeFolder(importRoot, relativeFolder);
            }

            destinationPath = relativeFolder;
        }

        object? opened = null;
        object? moved = null;
        try
        {
            opened = Call(session, "OpenSharedItem", msgPath)
                ?? throw new InvalidDataException("Outlook에서 변환된 MSG를 열 수 없습니다.");
            if (Get(opened, "Class") is not int itemClass || itemClass != OlMail)
            {
                throw new InvalidDataException("Outlook 일반 메일 항목으로 열리지 않는 메시지입니다.");
            }

            // Saving the open MSG first can create an item in the default mailbox.
            moved = Call(opened, "Move", destination ?? importRoot)
                ?? throw new InvalidOperationException("Outlook에서 PST로 이동한 메일을 반환하지 않았습니다.");
        }
        finally
        {
            if (moved is null && opened is not null)
            {
                CloseWithoutSaving(opened);
            }

            Release(moved);
            Release(opened);
        }
    }

    public void Detach()
    {
        ReleaseFolders();
        if (workingPath is null)
        {
            return;
        }

        storeRoot ??= FindStoreRoot(workingPath);
        if (storeRoot is not null)
        {
            Call(session, "RemoveStore", storeRoot);
        }

        Release(storeRoot);
        storeRoot = null;
        workingPath = null;
    }

    public void AttachExisting(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("연결할 PST를 찾을 수 없습니다.", path);
        }

        Call(session, "AddStoreEx", path, OlStoreUnicode);
    }

    public void Dispose()
    {
        ReleaseFolders();
        Release(storeRoot);
        Release(session);
        Release(application);
    }

    private object? FindStoreRoot(string path)
    {
        object stores = Get(session, "Stores")
            ?? throw new InvalidOperationException("Outlook 데이터 파일 목록을 열 수 없습니다.");
        try
        {
            int count = (int)(Get(stores, "Count") ?? 0);
            for (int index = 1; index <= count; index++)
            {
                object? store = Get(stores, "Item", index);
                if (store is null)
                {
                    continue;
                }

                try
                {
                    if (Get(store, "FilePath") is string storePath &&
                        string.Equals(storePath, path, StringComparison.OrdinalIgnoreCase))
                    {
                        return Call(store, "GetRootFolder");
                    }
                }
                finally
                {
                    Release(store);
                }
            }

            return null;
        }
        finally
        {
            Release(stores);
        }
    }

    private static object OpenRelativeFolder(object root, string relativeFolder)
    {
        object? child = null;
        try
        {
            foreach (string name in relativeFolder.Split(Path.DirectorySeparatorChar))
            {
                object next = OpenOrCreateChild(child ?? root, name);
                Release(child);
                child = next;
            }

            object result = child ?? throw new InvalidOperationException("가져오기 폴더를 찾을 수 없습니다.");
            child = null;
            return result;
        }
        finally
        {
            Release(child);
        }
    }

    private static object OpenOrCreateChild(object parent, string name)
    {
        object folders = Get(parent, "Folders")
            ?? throw new InvalidOperationException("PST 하위 폴더 목록을 열 수 없습니다.");
        try
        {
            int count = (int)(Get(folders, "Count") ?? 0);
            for (int index = 1; index <= count; index++)
            {
                object? child = Get(folders, "Item", index);
                if (child is null)
                {
                    continue;
                }

                try
                {
                    if (string.Equals(Get(child, "Name") as string, name, StringComparison.OrdinalIgnoreCase))
                    {
                        object result = child;
                        child = null;
                        return result;
                    }
                }
                finally
                {
                    Release(child);
                }
            }

            return Call(folders, "Add", name, OlFolderInbox)
                ?? throw new InvalidOperationException($"PST 폴더를 만들 수 없습니다. ({name})");
        }
        finally
        {
            Release(folders);
        }
    }

    private void ReleaseFolders()
    {
        Release(destination);
        destination = null;
        destinationPath = null;
        Release(importRoot);
        importRoot = null;
    }

    private static void CloseWithoutSaving(object item)
    {
        try
        {
            Call(item, "Close", OlDiscard);
        }
        catch (COMException)
        {
            // An already-closed item still needs its COM reference released.
        }
    }

    private static object? Get(object target, string name, params object[] arguments)
    {
        return target.GetType().InvokeMember(name, BindingFlags.GetProperty | BindingFlags.DoNotWrapExceptions,
            null, target, arguments, CultureInfo.InvariantCulture);
    }

    private static object? Call(object target, string name, params object[] arguments)
    {
        return target.GetType().InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.DoNotWrapExceptions,
            null, target, arguments, CultureInfo.InvariantCulture);
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            Marshal.ReleaseComObject(value);
        }
    }
}
