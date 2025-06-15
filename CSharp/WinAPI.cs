using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace MagicExtractor
{
    public class WinAPI
    {
        [Flags]
        private enum FileOperationFlags: uint
        {
            FOF_SILENT = 0x00000004,
            FOF_NOCONFIRMATION = 0x00000010,
            FOF_ALLOWUNDO = 0x00000040,
        }

        [Flags]
        private enum FileOperationType : uint
        {
            FO_DELETE = 0x00000003,
        }


        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct SHFILEOPSTRUCT
        {
            public IntPtr hwnd;
            public FileOperationType wFunc;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pFrom;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string? pTo;
            public FileOperationFlags fFlags;
            [MarshalAs(UnmanagedType.Bool)]
            public bool fAnyOperationsAborted;
            public IntPtr hNameMappings;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string? lpszProgressTitle;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHFileOperation(ref SHFILEOPSTRUCT FileOp);

        public bool SendFileToRecycleBin(string FilePath)//FIXME: this function not work on windows service
        {
            if (File.Exists(FilePath))
            {
                SHFILEOPSTRUCT ShStruct = new SHFILEOPSTRUCT
                {
                    hwnd = IntPtr.Zero,
                    wFunc = FileOperationType.FO_DELETE,
                    pFrom = FilePath + '\0' + '\0',
                    pTo = null,
                    fFlags = FileOperationFlags.FOF_ALLOWUNDO | FileOperationFlags.FOF_NOCONFIRMATION | FileOperationFlags.FOF_SILENT,
                    fAnyOperationsAborted = false,
                    hNameMappings = IntPtr.Zero,
                    lpszProgressTitle = null
                };

                if (SHFileOperation(ref ShStruct) != 0) { return false; }
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
