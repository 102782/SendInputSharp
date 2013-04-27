using System;
using System.Runtime.InteropServices;

namespace SendInputSharp
{
    [StructLayout(LayoutKind.Explicit)]
    internal struct MOUSEINPUT
    {
        [FieldOffset(0)]
        public int dx;
        [FieldOffset(4)]
        public int dy;
        [FieldOffset(8)]
        public int mouseData;
        [FieldOffset(12)]
        public int dwFlags;
        [FieldOffset(16)]
        public int time;
        [FieldOffset(20)]
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Explicit)]
    internal struct KEYBDINPUT
    {
        [FieldOffset(0)]
        public ushort wVk;
        [FieldOffset(2)]
        public ushort wScan;
        [FieldOffset(4)]
        public int dwFlags;
        [FieldOffset(8)]
        public int time;
        [FieldOffset(12)]
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Explicit)]
    internal struct HARDWAREINPUT
    {
        [FieldOffset(0)]
        public int uMsg;
        [FieldOffset(4)]
        public ushort wParamL;
        [FieldOffset(6)]
        public ushort wParamH;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct INPUT
    {
        internal uint type;
        internal InputUnion U;
        internal static int Size
        {
            get { return Marshal.SizeOf(typeof(INPUT)); }
        }
    }

    [StructLayout(LayoutKind.Explicit)]
    internal struct InputUnion
    {
        [FieldOffset(0)]
        internal MOUSEINPUT mi;
        [FieldOffset(0)]
        internal KEYBDINPUT ki;
        [FieldOffset(0)]
        internal HARDWAREINPUT hi;
    }
}