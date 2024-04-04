namespace cs_console_app
{
    using System;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.ComTypes;

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