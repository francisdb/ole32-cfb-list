using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

using HRESULT = System.Int32;
public class Com
{
    public const HRESULT S_OK = 0;
}

[ComImport]
[Guid("0000000d-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IEnumSTATSTG
{
    // The user needs to allocate an STATSTG array whose size is celt.
    [PreserveSig]
    uint Next(
        uint celt,
        [MarshalAs(UnmanagedType.LPArray), Out]
            System.Runtime.InteropServices.ComTypes.STATSTG[] rgelt,
        out uint pceltFetched
    );

    void Skip(uint celt);

    void Reset();

    [return: MarshalAs(UnmanagedType.Interface)]
    IEnumSTATSTG Clone();
}

[ComImport]
[Guid("0000000b-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IStorage
{
    void CreateStream(
        /* [string][in] */ string pwcsName,
        /* [in] */ STGM grfMode,
        /* [in] */ uint reserved1,
        /* [in] */ uint reserved2,
        /* [out] */ out IStream ppstm);

    void OpenStream(
        /* [string][in] */ string pwcsName,
        /* [unique][in] */ IntPtr reserved1,
        /* [in] */ STGM grfMode,
        /* [in] */ uint reserved2,
        /* [out] */ out IStream ppstm);

    void CreateStorage(
        /* [string][in] */ string pwcsName,
        /* [in] */ STGM grfMode,
        /* [in] */ uint reserved1,
        /* [in] */ uint reserved2,
        /* [out] */ out IStorage ppstg);

    void OpenStorage(
        /* [string][unique][in] */ string pwcsName,
        /* [unique][in] */ IStorage? pstgPriority,
        /* [in] */ STGM grfMode,
        /* [unique][in] */ IntPtr snbExclude,
        /* [in] */ uint reserved,
        /* [out] */ out IStorage ppstg);

    void CopyTo(
        /* [in] */ uint ciidExclude,
        /* [size_is][unique][in] */ Guid rgiidExclude, // should this be an array?
        /* [unique][in] */ IntPtr snbExclude,
        /* [unique][in] */ IStorage pstgDest);

    void MoveElementTo(
        /* [string][in] */ string pwcsName,
        /* [unique][in] */ IStorage pstgDest,
        /* [string][in] */ string pwcsNewName,
        /* [in] */ uint grfFlags);

    void Commit(
        /* [in] */ uint grfCommitFlags);

    void Revert();

    void EnumElements(
        /* [in] */ uint reserved1,
        /* [size_is][unique][in] */ IntPtr reserved2,
        /* [in] */ uint reserved3,
        /* [out] */ out IEnumSTATSTG ppenum);

    void DestroyElement(
        /* [string][in] */ string pwcsName);

    void RenameElement(
        /* [string][in] */ string pwcsOldName,
        /* [string][in] */ string pwcsNewName);

    void SetElementTimes(
        /* [string][unique][in] */ string pwcsName,
        /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pctime,
        /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME patime,
        /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pmtime);

    void SetClass(
        /* [in] */ Guid clsid);

    void SetStateBits(
        /* [in] */ uint grfStateBits,
        /* [in] */ uint grfMask);

    void Stat(
        /* [out] */ out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg,
        /* [in] */ uint grfStatFlag);

}

[Flags]
public enum STGM : int
{
    /**  */
    DIRECT = 0x00000000,
    TRANSACTED = 0x00010000,
    SIMPLE = 0x08000000,
    READ = 0x00000000,
    WRITE = 0x00000001,
    READWRITE = 0x00000002,
    SHARE_DENY_NONE = 0x00000040,
    SHARE_DENY_READ = 0x00000030,
    SHARE_DENY_WRITE = 0x00000020,
    SHARE_EXCLUSIVE = 0x00000010,
    PRIORITY = 0x00040000,
    DELETEONRELEASE = 0x04000000,
    NOSCRATCH = 0x00100000,
    CREATE = 0x00001000,
    CONVERT = 0x00020000,
    FAILIFTHERE = 0x00000000,
    NOSNAPSHOT = 0x00200000,
    DIRECT_SWMR = 0x00400000,
}

public enum STATFLAG : uint
{
    STATFLAG_DEFAULT = 0,
    STATFLAG_NONAME = 1,
    STATFLAG_NOOPEN = 2
}

public enum STGTY : int
{
    STGTY_STORAGE = 1,
    STGTY_STREAM = 2,
    STGTY_LOCKBYTES = 3,
    STGTY_PROPERTY = 4
}

public enum STGFMT : int
{
    STGFMT_STORAGE = 0,
    STGFMT_FILE = 3,
    STGFMT_ANY = 4,
    STGFMT_DOCFILE = 5
}

[StructLayout(LayoutKind.Sequential)]
public struct STGOPTIONS
{
    public ushort usVersion;
    public ushort reserved;
    public uint ulSectorSize;
}

public static class Ole32
{
    [DllImport("ole32.dll")]
    public static extern int StgIsStorageFile(
        [MarshalAs(UnmanagedType.LPWStr)] string pwcsName);

    // https://learn.microsoft.com/en-us/windows/win32/api/coml2api/nf-coml2api-stgopenstorage
    [DllImport("ole32.dll")]
    public static extern int StgOpenStorage(
        [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
        IStorage? pstgPriority,
        STGM grfMode,
        IntPtr snbExclude,
        uint reserved,
        out IStorage ppstgOpen);

    // https://learn.microsoft.com/en-us/windows/win32/api/coml2api/nf-coml2api-stgcreatestorageex
    [DllImport("ole32.dll")]
    public static extern int StgCreateStorageEx(
        [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
        STGM grfMode,
        STGFMT stgfmt,
        uint grfAttrs,
        ref STGOPTIONS pStgOptions,
        IntPtr pSecurityDescriptor,
        ref Guid riid,
        out IStorage ppObjectOpen);
}

public class DisposableIStream : IDisposable
{
    public IStream Stream { get; private set; }

    public DisposableIStream(IStream stream)
    {
        Stream = stream;
    }

    public DisposableIStream(IStorage storage, string pwcsName, STGM grfMode)
    {
        storage.OpenStream(pwcsName, IntPtr.Zero, grfMode, 0, out IStream stream);
        Stream = stream;
    }

    public int Read(byte[] buffer, int length)
    {
        IntPtr pcbRead = Marshal.AllocHGlobal(sizeof(int));
        Stream.Read(buffer, length, pcbRead);
        int bytesRead = Marshal.ReadInt32(pcbRead);
        Marshal.FreeHGlobal(pcbRead);
        return bytesRead;
    }

    public STATSTG Stat(int grfStatFlag)
    {
        Stream.Stat(out STATSTG statstg, grfStatFlag);
        return statstg;
    }

    public void Write(byte[] buffer, int length)
    {
        IntPtr pcbWritten = Marshal.AllocHGlobal(sizeof(int));
        Stream.Write(buffer, length, pcbWritten);
        Marshal.FreeHGlobal(pcbWritten);
    }

    public void Dispose()
    {
        if (Stream != null)
        {
            Marshal.ReleaseComObject(Stream);
            //Stream = null;
        }
    }
}

public class DisposableIStorage : IDisposable
{

    public static DisposableIStorage CreateStorage(string pwcsName, STGM grfMode)
    {
        // guid ref for V4
        Guid IID_IPropertySetStorage = new Guid("0000013A-0000-0000-C000-000000000046");

        STGOPTIONS stg;
        stg.usVersion = 1;
        stg.reserved = 0;
        stg.ulSectorSize = 4096;

        HRESULT hr = Ole32.StgCreateStorageEx(pwcsName, grfMode, STGFMT.STGFMT_DOCFILE, 0, ref stg, IntPtr.Zero, ref IID_IPropertySetStorage, out IStorage storage);
        if (hr != Com.S_OK)
        {
            Exception? ex = Marshal.GetExceptionForHR(hr);
            throw new Exception("Error while creating file: " + (ex?.Message));
        }
        return new DisposableIStorage(storage);
    }

    public IStorage Storage { get; private set; }

    private DisposableIStorage(IStorage storage)
    {
        Storage = storage;
    }

    public DisposableIStorage(string pwcsName, IStorage? pstgPriority, STGM grfMode, IntPtr snbExclude)
    {
        // do we want to do a if ( Ole32.StgIsStorageFile(filename) == 0) first?
        HRESULT hr = Ole32.StgOpenStorage(pwcsName, null, grfMode, snbExclude, 0, out IStorage storage);
        if (hr != Com.S_OK)
        {
            Exception? ex = Marshal.GetExceptionForHR(hr);
            throw new Exception("Error while opening file: " + (ex?.Message));
        }
        Storage = storage;
    }

    public DisposableIStorage OpenStorage(string pwcsName, IStorage? pstgPriority, STGM grfMode, IntPtr snbExclude)
    {
        Storage.OpenStorage(pwcsName, pstgPriority, grfMode, snbExclude, 0, out IStorage subStorage);
        return new DisposableIStorage(subStorage);
    }

    public DisposableIStream OpenStream(string pwcsName, IntPtr reserved1, STGM grfMode)
    {
        Storage.OpenStream(pwcsName, reserved1, grfMode, 0, out IStream stream);
        return new DisposableIStream(stream);
    }

    public DisposableIStream CreateStream(string pwcsName, STGM grfMode)
    {
        Storage.CreateStream(pwcsName, grfMode, 0, 0, out IStream stream);
        return new DisposableIStream(stream);
    }

    public IEnumerator<STATSTG> EnumElements()
    {
        Storage.EnumElements(0, IntPtr.Zero, 0, out IEnumSTATSTG enumStatStg);
        STATSTG[] statStg = new STATSTG[1];
        uint fetched = 0;
        while (enumStatStg.Next(1, statStg, out fetched) == 0 && fetched > 0)
        {
            yield return statStg[0];
        }
        Marshal.ReleaseComObject(enumStatStg);
    }

    public void Dispose()
    {
        if (Storage != null)
        {
            Marshal.ReleaseComObject(Storage);
            //Storage = null;
        }
    }
}