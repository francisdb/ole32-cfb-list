namespace cs_console_app
{
    using System;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.ComTypes;
    using System.Text.RegularExpressions;

    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Current directory: " + Environment.CurrentDirectory);
            string filename_minimal_ole32 = @"../../../../minimal.ole32.cfb";
            string filename_minimal_rust = @"../../../../minimal.rust.cfb";
            Console.WriteLine("================= Diff ========================");
            Diff(filename_minimal_ole32, filename_minimal_rust, new List<string> { });
            // Console.WriteLine("================= Original ====================");
            // ListPaths(filename_original);
            // Console.WriteLine("================= Copy ========================");
            // ListPaths(filename_copy);

            // open stream root\stream31
            // using (var storage = new DisposableIStorage(filename_minimal_rust, null, STGM.DIRECT | STGM.READ | STGM.SHARE_EXCLUSIVE, IntPtr.Zero))
            // {
            //     try{
            //     using (var stream = storage.OpenStream("stream31", IntPtr.Zero, STGM.READ | STGM.SHARE_EXCLUSIVE))
            //     {
            //         STATSTG statstg = stream.Stat((int)STATFLAG.STATFLAG_DEFAULT);
            //         byte[] buffer = new byte[statstg.cbSize];
            //         stream.Read(buffer, buffer.Length);
            //         Console.WriteLine("Stream content: " + buffer);
            //     }
            //     }catch(Exception e){
            //         Console.WriteLine("Error: " + e.Message);
            //     }
            // }

            // // create a new cfg file with 32 empty streams
            // using (var storage = DisposableIStorage.CreateStorage(filename_minimal_ole32, STGM.READWRITE | STGM.CREATE | STGM.SHARE_EXCLUSIVE | STGM.DIRECT))
            // {
            //     for (int i = 0; i < 32; i++)
            //     {
            //         using (var stream = storage.CreateStream($"stream{i}", STGM.READWRITE | STGM.CREATE | STGM.SHARE_EXCLUSIVE))
            //         {
            //             byte[] buffer = new byte[0];
            //             stream.Write(buffer, buffer.Length);
            //         }
            //     }
            // }
            // ListPaths(filename_minimal_ole32);
        }

        /**
            Diff two storage files
            Uses the first one as the reference and compares it with the second one
        */
        private static void Diff(string filename1, string filename2, List<string> ignoreInPaths)
        {
            using (var storage1 = new DisposableIStorage(filename1, null, STGM.DIRECT | STGM.READ | STGM.SHARE_EXCLUSIVE, IntPtr.Zero))
            {
                using (var storage2 = new DisposableIStorage(filename2, null, STGM.DIRECT | STGM.READ | STGM.SHARE_EXCLUSIVE, IntPtr.Zero))
                {
                    DoDiff(storage1, storage2, "root", ignoreInPaths);

                    List<string> nodes1 = EnumerateNodes(storage1, "root");
                    List<string> nodes2 = EnumerateNodes(storage2, "root");
                    foreach (string node in nodes2)
                    {
                        if (!nodes1.Contains(node))
                        {
                            Console.WriteLine($"⚠️ {node} Extra entry in second file");
                        }
                    }
                }
            }
        }

        private static int DoDiff(DisposableIStorage storage1, DisposableIStorage storage2, string currentPath, List<string> ignoreInPaths, int count = 0)
        {
            Console.WriteLine($"📁 {count} {currentPath}");
            IEnumerator<STATSTG> enumerator1 = storage1.EnumElements();
            while (enumerator1.MoveNext())
            {
                STATSTG statsg1 = enumerator1.Current;
                string strNode1 = statsg1.pwcsName;

                string fullPath = Path.Combine(currentPath, strNode1);
                if (ignoreInPaths.Any(path => fullPath.Contains(path)))
                {
                    continue;
                }

                count++;

                if (statsg1.type == (int)STGTY.STGTY_STREAM)
                {
                    try
                    {
                        using (var stream1 = storage1.OpenStream(strNode1, IntPtr.Zero, STGM.READ | STGM.SHARE_EXCLUSIVE))
                        {
                            using (var stream2 = storage2.OpenStream(strNode1, IntPtr.Zero, STGM.READ | STGM.SHARE_EXCLUSIVE))
                            {
                                StreamDiffResult result = DiffStream(fullPath, stream1, stream2);
                                switch (result)
                                {
                                    case StreamDiffResult.Same:
                                        Console.WriteLine($"✅ {count} {fullPath} Same");
                                        break;
                                    case StreamDiffResult.DifferentSize:
                                        Console.WriteLine($"☑️ {count} {fullPath} Different size");
                                        break;
                                    case StreamDiffResult.DifferentContent:
                                        Console.WriteLine($"❌ {count} {fullPath} Different content");
                                        break;
                                }
                            }
                        }
                    }
                    catch (COMException e)
                    {
                        if (e.ErrorCode == unchecked((int)0x80030002)) // STG_E_FILENOTFOUND
                        {
                            Console.WriteLine($"➖ {count} {fullPath} Not found");
                        }
                        else
                        {
                            Console.WriteLine($"❌ {count} {fullPath} Error: {e.Message}");
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"❌ {count} {fullPath} Error: {e.Message}");
                    }
                }

                if (statsg1.type == (int)STGTY.STGTY_STORAGE)
                {
                    using (var subStorage1 = storage1.OpenStorage(strNode1, null, STGM.READ | STGM.SHARE_EXCLUSIVE, IntPtr.Zero))
                    {
                        using (var subStorage2 = storage2.OpenStorage(strNode1, null, STGM.READ | STGM.SHARE_EXCLUSIVE, IntPtr.Zero))
                        {
                            count = DoDiff(subStorage1, subStorage2, fullPath, ignoreInPaths, count);
                        }
                    }
                }
            }
            return count;
        }

        // create an enum for the stream diff result
        private enum StreamDiffResult
        {
            Same,
            DifferentSize,
            DifferentContent
        }

        private static StreamDiffResult DiffStream(string fullPath, DisposableIStream stream1, DisposableIStream stream2)
        {
            STATSTG statstg1 = stream1.Stat((int)STATFLAG.STATFLAG_DEFAULT);
            STATSTG statstg2 = stream2.Stat((int)STATFLAG.STATFLAG_DEFAULT);
            if (statstg1.cbSize != statstg2.cbSize)
            {
                return StreamDiffResult.DifferentSize;
            }
            else
            {
                byte[] buffer1 = new byte[statstg1.cbSize];
                byte[] buffer2 = new byte[statstg2.cbSize];
                stream1.Read(buffer1, buffer1.Length);
                stream2.Read(buffer2, buffer2.Length);
                bool equal = true;
                for (int i = 0; i < buffer1.Length; i++)
                {
                    if (buffer1[i] != buffer2[i])
                    {
                        equal = false;
                        break;
                    }
                }
                if (!equal)
                {
                    return StreamDiffResult.DifferentContent;
                }
                else
                {
                    return StreamDiffResult.Same;
                }
            }
        }

        private static void ListPaths(string filename)
        {
            Console.WriteLine("Opening storage at " + filename);
            using (var storage = new DisposableIStorage(filename, null, STGM.DIRECT | STGM.READ | STGM.SHARE_EXCLUSIVE, IntPtr.Zero))
            {
                List<string> nodes = EnumerateNodes(storage, "root");
                foreach (string node in nodes)
                {
                    Console.WriteLine(node);
                }
            }
        }

        static List<string> EnumerateNodes(DisposableIStorage storage, string currentPath)
        {
            List<string> nodes = new List<string>();
            IEnumerator<STATSTG> enumerator = storage.EnumElements();
            while (enumerator.MoveNext())
            {
                STATSTG statsg = enumerator.Current;
                string strNode = statsg.pwcsName;
                string fullPath = Path.Combine(currentPath, strNode);
                nodes.Add(fullPath);

                if (statsg.type == (int)STGTY.STGTY_STORAGE)
                {
                    using (var subStorage = storage.OpenStorage(strNode, null, STGM.READ | STGM.SHARE_EXCLUSIVE, IntPtr.Zero))
                    {
                        nodes.AddRange(EnumerateNodes(subStorage, fullPath));
                    }
                }
            }
            return nodes;
        }
    }
}