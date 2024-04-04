namespace cs_console_app
{
    using System;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.ComTypes;

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
    interface IStorage
    {
        void CreateStream(
            /* [string][in] */ string pwcsName,
            /* [in] */ uint grfMode,
            /* [in] */ uint reserved1,
            /* [in] */ uint reserved2,
            /* [out] */ out IStream ppstm);

        void OpenStream(
            /* [string][in] */ string pwcsName,
            /* [unique][in] */ IntPtr reserved1,
            /* [in] */ uint grfMode,
            /* [in] */ uint reserved2,
            /* [out] */ out IStream ppstm);

        void CreateStorage(
            /* [string][in] */ string pwcsName,
            /* [in] */ uint grfMode,
            /* [in] */ uint reserved1,
            /* [in] */ uint reserved2,
            /* [out] */ out IStorage ppstg);

        void OpenStorage(
            /* [string][unique][in] */ string pwcsName,
            /* [unique][in] */ IStorage pstgPriority,
            /* [in] */ uint grfMode,
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

    class Program
    {
        [DllImport("ole32.dll")]
        private static extern int StgIsStorageFile(
            [MarshalAs(UnmanagedType.LPWStr)] string pwcsName);

        [DllImport("ole32.dll")]
        static extern int StgOpenStorage(
            [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
            IStorage pstgPriority,
            STGM grfMode,
            IntPtr snbExclude,
            uint reserved,
            out IStorage ppstgOpen);

        static void Main(string[] args)
        {
            string filename_original = @"c:\Users\franc\RustroverProjects\aztec-quest\aztecquest-dev.vpx";
            string filename_copy = @"c:\Users\franc\RustroverProjects\aztec-quest\aztecquest-copy.vpx";
            Console.WriteLine("================= Original ====================");
            ListVPXContents(filename_original);
            Console.WriteLine("================= Copy ========================");
            ListVPXContents(filename_copy);
            Console.WriteLine("================= Original name ===============");
            ReadTableInfoTableName(filename_original);
            Console.WriteLine("================= Copy name ===================");
            ReadTableInfoTableName(filename_copy);
        }

        // read root\TableInfo\TableName
        private static void ReadTableInfoTableName(string filename)
        {
            IStorage storage = null;
            if (StgOpenStorage(
                filename,
                null,
                STGM.DIRECT | STGM.READ | STGM.SHARE_EXCLUSIVE,
                IntPtr.Zero,
                0,
                out storage) != 0)
            {
                Console.WriteLine("Failed to open storage at " + filename);
                return;
            }

            IStorage tableInfoStorage;
            storage.OpenStorage("TableInfo", null, (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE), IntPtr.Zero, 0, out tableInfoStorage);

            IStream tableNameStream;
            tableInfoStorage.OpenStream("TableName", IntPtr.Zero, (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE), 0, out tableNameStream);

            System.Runtime.InteropServices.ComTypes.STATSTG statstg;
            tableNameStream.Stat(out statstg, (int)STATFLAG.STATFLAG_DEFAULT);
            IntPtr pcbRead = IntPtr.Zero;
            byte[] buffer = new byte[statstg.cbSize];
            tableNameStream.Read(buffer, buffer.Length, pcbRead);
            Marshal.ReleaseComObject(tableNameStream);
            Marshal.ReleaseComObject(tableInfoStorage);
            Marshal.ReleaseComObject(storage);

            // the string is encoded with 2 bytes per character
            Console.WriteLine(System.Text.Encoding.Unicode.GetString(buffer));
        }

        private static void ListVPXContents(string filename)
        {
            if (StgIsStorageFile(filename) == 0)
            {
                Console.WriteLine("Opening storage at " + filename);
                IStorage storage = null;
                if (StgOpenStorage(
                    filename,
                    null,
                    STGM.DIRECT | STGM.READ | STGM.SHARE_EXCLUSIVE,
                    IntPtr.Zero,
                    0,
                    out storage) == 0)
                {
                    EnumerateNodes(storage, "root");
                    Marshal.ReleaseComObject(storage);
                }
                else
                {
                    Console.WriteLine("Failed to open storage at " + filename);
                }
            }
            else
            {
                Console.WriteLine("Not a storage file: " + filename);
            }
        }

        static void EnumerateNodes(IStorage storage, string currentPath)
        {
            IEnumSTATSTG pIEnumStatStg = null;
            storage.EnumElements(0, IntPtr.Zero, 0, out pIEnumStatStg);

            System.Runtime.InteropServices.ComTypes.STATSTG[] regelt = new System.Runtime.InteropServices.ComTypes.STATSTG[1];
            uint fetched = 0;

            while (pIEnumStatStg.Next(1, regelt, out fetched) == 0 && fetched > 0)
            {
                string strNode = regelt[0].pwcsName;
                string fullPath = Path.Combine(currentPath, strNode);
                Console.WriteLine(fullPath);

                if (regelt[0].type == (int)STGTY.STGTY_STORAGE)
                {
                    IStorage subStorage;
                    storage.OpenStorage(strNode, null, (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE), IntPtr.Zero, 0, out subStorage);
                    EnumerateNodes(subStorage, fullPath);
                    Marshal.ReleaseComObject(subStorage);
                }
            }
            Marshal.ReleaseComObject(pIEnumStatStg);
        }
    }
}