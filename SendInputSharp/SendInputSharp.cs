using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace SendInputSharp
{
    /// <summary>
    /// KEYEVENTF定数
    /// </summary>
    [Flags]
    public enum KEYEVENTFLAG : uint
    {
        KEYDOWN     = 0x00000000,
        EXTENDEDKEY = 0x00000001,
        KEYUP       = 0x00000002,
        UNICODE     = 0x00000004,
        SCANCODE    = 0x00000008
    }

    internal enum INPUTTYPE : uint
    {
        MOUSE = 0,
        KEYBOARD = 1,
        HARDWARE = 2
    }

    /// <summary>
    /// 仮想キーコード
    /// </summary>
    public enum VIRTUALKEYCODE : short
    {
        /// <summary>
        /// BREAK(Control+Pause) key(ExtendedKey)
        /// </summary>
        VK_CANCEL = 0x0003,
        /// <summary>
        /// BACKSPACE key
        /// </summary>
        VK_BACK = 0x0008,
        VK_TAB = 0x0009,        //TAB key
        VK_CLEAR = 0x000C,
        VK_RETURN = 0x000D,     //ENTER key
        VK_SHIFT = 0x0010,      //SHIFT key
        /// <summary>
        /// CTRL key
        /// </summary>
        VK_CONTROL = 0x0011,
        /// <summary>
        /// ALT key
        /// </summary>
        VK_MENU = 0x0012,
        VK_PAUSE = 0x0013,      //PAUSE key
        /// <summary>
        /// CAPS LOCK key
        /// </summary>
        VK_CAPITAL = 0x0014,
        /// <summary>
        /// IME かな mode
        /// </summary>
        VK_KANA = 0x0015,
        VK_JUNJA = 0x0017,
        VK_FINAL = 0x0018,
        /// <summary>
        /// IME 漢字 mode
        /// </summary>
        VK_KANJI = 0x0019,
        /// <summary>
        /// ESC key
        /// </summary>
        VK_ESCAPE = 0x001B,
        /// <summary>
        /// IME 変換 key
        /// </summary>
        VK_CONVERT = 0x001C,
        /// <summary>
        /// IME 無変換 key
        /// </summary>
        VK_NONCONVERT = 0x001D,
        VK_ACCEPT = 0x001E,
        VK_MODECHANGE = 0x001F,
        /// <summary>
        /// SPACEBAR
        /// </summary>
        VK_SPACE = 0x0020,
        /// <summary>
        /// PAGE UP key(ExtendedKey)
        /// </summary>
        VK_PRIOR = 0x0021,
        /// <summary>
        /// PAGE DOWN key(ExtendedKey)
        /// </summary>
        VK_NEXT = 0x0022,
        /// <summary>
        /// END key(ExtendedKey)
        /// </summary>
        VK_END = 0x0023,
        /// <summary>
        /// HOME key(ExtendedKey)
        /// </summary>
        VK_HOME = 0x0024,
        /// <summary>
        /// ← key(ExtendedKey)
        /// </summary>
        VK_LEFT = 0x0025,
        /// <summary>
        /// ↑ key(ExtendedKey)
        /// </summary>
        VK_UP = 0x0026,
        /// <summary>
        /// → key(ExtendedKey)
        /// </summary>
        VK_RIGHT = 0x0027,
        /// <summary>
        /// ↓ key(ExtendedKey)
        /// </summary>
        VK_DOWN = 0x0028,
        VK_SELECT = 0x0029,
        VK_PRINT = 0x002A,
        VK_EXECUTE = 0x002B,
        /// <summary>
        /// PRINT SCREEN key(ExtendedKey)
        /// </summary>
        VK_SNAPSHOT = 0x002C,
        /// <summary>
        /// INS key(ExtendedKey)
        /// </summary>
        VK_INSERT = 0x002D,
        /// <summary>
        /// DEL key(ExtendedKey)
        /// </summary>
        VK_DELETE = 0x002E,
        VK_HELP = 0x002F,
        VK_0 = 0x0030,          //0 key
        VK_1 = 0x0031,          //1 key
        VK_2 = 0x0032,          //2 key
        VK_3 = 0x0033,          //3 key
        VK_4 = 0x0034,          //4 key
        VK_5 = 0x0035,          //5 key
        VK_6 = 0x0036,          //6 key
        VK_7 = 0x0037,          //7 key
        VK_8 = 0x0038,          //8 key
        VK_9 = 0x0039,          //9 key
        VK_A = 0x0041,          //A key
        VK_B = 0x0042,          //B key
        VK_C = 0x0043,          //C key
        VK_D = 0x0044,          //D key
        VK_E = 0x0045,          //E key
        VK_F = 0x0046,          //F key
        VK_G = 0x0047,          //G key
        VK_H = 0x0048,          //H key
        VK_I = 0x0049,          //I key
        VK_J = 0x004A,          //J key
        VK_K = 0x004B,          //K key
        VK_L = 0x004C,          //L key
        VK_M = 0x004D,          //M key
        VK_N = 0x004E,          //N key
        VK_O = 0x004F,          //O key
        VK_P = 0x0050,          //P key
        VK_Q = 0x0051,          //Q key
        VK_R = 0x0052,          //R key
        VK_S = 0x0053,          //S key
        VK_T = 0x0054,          //T key
        VK_U = 0x0055,          //U key
        VK_V = 0x0056,          //V key
        VK_W = 0x0057,          //W key
        VK_X = 0x0058,          //X key
        VK_Y = 0x0059,          //Y key
        VK_Z = 0x005A,          //Z key
        /// <summary>
        /// Left Windows key
        /// </summary>
        VK_LWIN = 0x005B,
        /// <summary>
        /// Right Windows key
        /// </summary>
        VK_RWIN = 0x005C,
        /// <summary>
        /// Applications key
        /// </summary>
        VK_APPS = 0x005D,
        VK_NUMPAD0 = 0x0060,    //Numeric keypad 0 key
        VK_NUMPAD1 = 0x0061,    //Numeric keypad 1 key
        VK_NUMPAD2 = 0x0062,    //Numeric keypad 2 key
        VK_NUMPAD3 = 0x0063,    //Numeric keypad 3 key
        VK_NUMPAD4 = 0x0064,    //Numeric keypad 4 key
        VK_NUMPAD5 = 0x0065,    //Numeric keypad 5 key
        VK_NUMPAD6 = 0x0066,    //Numeric keypad 6 key
        VK_NUMPAD7 = 0x0067,    //Numeric keypad 7 key
        VK_NUMPAD8 = 0x0068,    //Numeric keypad 8 key
        VK_NUMPAD9 = 0x0069,    //Numeric keypad 9 key
        /// <summary>
        /// * key
        /// </summary>
        VK_MULTIPLY = 0x006A,
        /// <summary>
        /// + key
        /// </summary>
        VK_ADD = 0x006B,
        VK_SEPERATOR = 0x006C,
        /// <summary>
        /// - key
        /// </summary>
        VK_SUBTRACT = 0x006D,
        /// <summary>
        /// テンキーの . key
        /// </summary>
        VK_DECIMAL = 0x006E,
        /// <summary>
        /// /key(ExtendedKey)
        /// </summary>
        VK_DIVIDE = 0x006F,
        VK_F1 = 0x0070,         //F1 key
        VK_F2 = 0x0071,         //F2 key
        VK_F3 = 0x0072,         //F3 key
        VK_F4 = 0x0073,         //F4 key
        VK_F5 = 0x0074,         //F5 key
        VK_F6 = 0x0075,         //F6 key
        VK_F7 = 0x0076,         //F7 key
        VK_F8 = 0x0077,         //F8 key
        VK_F9 = 0x0078,         //F9 key
        VK_F10 = 0x0079,        //F10 key
        VK_F11 = 0x007A,        //F11 key
        VK_F12 = 0x007B,        //F12 key
        VK_F13 = 0x007C,        //F13 key
        VK_F14 = 0x007D,        //F14 key
        VK_F15 = 0x007E,        //F15 key
        VK_F16 = 0x007F,        //F16 key
        VK_F17 = 0x0080,        //F17 key
        VK_F18 = 0x0081,        //F18 key
        VK_F19 = 0x0082,        //F19 key
        VK_F20 = 0x0083,        //F20 key
        VK_F21 = 0x0084,        //F21 key
        VK_F22 = 0x0085,        //F22 key
        VK_F23 = 0x0086,        //F23 key
        VK_F24 = 0x0087,        //F24 key
        /// <summary>
        /// NUM LOCK key(ExtendedKey)
        /// </summary>
        VK_NUMLOCK = 0x0090,
        /// <summary>
        /// SCROLL LOCK key
        /// </summary>
        VK_SCROLL = 0x0091,
        /// <summary>
        /// Left SHIFT key
        /// </summary>
        VK_LSHIFT = 0x00A0,
        /// <summary>
        /// Right SHIFT key(ExtendedKey)
        /// </summary>
        VK_RSHIFT = 0x00A1,
        /// <summary>
        /// Left CONTROL key
        /// </summary>
        VK_LCONTROL = 0x00A2,
        /// <summary>
        /// Right CONTROL key(ExtendedKey)
        /// </summary>
        VK_RCONTROL = 0x00A3,
        /// <summary>
        /// Left MENU key
        /// </summary>
        VK_LMENU = 0x00A4,
        /// <summary>
        /// Right MENU key(ExtendedKey)
        /// </summary>
        VK_RMENU = 0x00A5,
        VK_BROWSER_BACK = 0x00A6,       //Browser Back key
        VK_BROWSER_FORWARD = 0x00A7,    //Browser Forward key
        VK_BROWSER_REFRESH = 0x00A8,    //Browser Refresh key
        VK_BROWSER_STOP = 0x00A9,       //Browser Stop key
        VK_BROWSER_SEARCH = 0x00AA,     //Browser Search key
        VK_BROWSER_FAVORITES = 0x00AB,  //Browser Favorites key
        VK_BROWSER_HOME = 0x00A6,       //Browser Start and Home key
        VK_VOLUME_MUTE = 0x00AD,        //Volume Mute key
        VK_VOLUME_DOWN = 0x00AE,        //Volume Down key
        VK_VOLUME_UP = 0x00AF,          //Volume Up key
        VK_MEDIA_NEXT_TRACK = 0x00B0,   //Next Track key
        VK_MEDIA_PREV_TRACK = 0x00B1,   //Previous Track key
        VK_MEDIA_STOP = 0x00B2,         //Stop Media key
        VK_MEDIA_PLAY_PAUSE = 0x00B3,   //Play/Pause Media key
        /// <summary>
        /// Start Mail key
        /// </summary>
        VK_LAUNCH_MAIL = 0x00B4,
        /// <summary>
        /// Select Media Key
        /// </summary>
        VK_LAUNCH_MEDIA_SELECT = 0x00B5,
        /// <summary>
        /// Start Application 1 key
        /// </summary>
        VK_LAUNCH_APP1 = 0x00B6,
        /// <summary>
        /// Start Application 2 key
        /// </summary>
        VK_LAUNCH_APP2 = 0x00B7,
        /// <summary>
        /// : *  key
        /// </summary>
        VK_OEM_1 = 0x00BA,
        /// <summary>
        /// ; + key
        /// </summary>
        VK_OEM_PLUS = 0x00BB,
        /// <summary>
        /// , < key
        /// </summary>
        VK_OEM_COMMA = 0x00BC,
        /// <summary>
        /// - = key
        /// </summary>
        VK_OEM_MINUS = 0x00BD,
        /// <summary>
        /// . > key
        /// </summary>
        VK_OEM_PERIOD = 0x00BE,
        /// <summary>
        /// / ? key
        /// </summary>
        VK_OEM_2 = 0x00BF,
        /// <summary>
        /// @ ` key
        /// </summary>
        VK_OEM_3 = 0x00C0,
        /// <summary>
        /// [ { key
        /// </summary>
        VK_OEM_4 = 0x00DB,
        /// <summary>
        /// \ | key
        /// </summary>
        VK_OEM_5 = 0x00DC,
        /// <summary>
        /// ] } key
        /// </summary>
        VK_OEM_6 = 0x00DD,
        /// <summary>
        /// ^ ~ key
        /// </summary>
        VK_OEM_7 = 0x00DE,
        VK_OEM_8 = 0x00DF,
        VK_PROCESSKEY = 0x00E5,
        /// <summary>
        /// 英数
        /// </summary>
        VK_OEM_ATTN = 0x00F0,
        /// <summary>
        /// \ _ key
        /// </summary>
        VK_OEM_102 = 0x00E2,
        /// <summary>
        /// カタカナひらがな
        /// </summary>
        VK_OEM_COPY = 0x00F2,
        /// <summary>
        /// 全角/半角
        /// </summary>
        VK_OEM_AUTO = 0x00F3,
        /// <summary>
        /// 全角/半角
        /// </summary>
        VK_OEM_ENLW = 0x00F4,
        /// <summary>
        /// ローマ字
        /// </summary>
        VK_OEM_BACKTAB = 0x00F5,
        VK_PACKET = 0x00E7,
        VK_ATTN = 0x00F6,
        VK_CRSEL = 0x00F7,
        VK_EXSEL = 0x00F8,
        VK_EREOF = 0x00F9,
        VK_PLAY = 0x00FA,
        VK_ZOOM = 0x00FB,
        VK_NONAME = 0x00FC,
        VK_PA1 = 0x00FD,
        VK_OEM_CLEAR = 0x00FE,
    }

    internal enum SCANCODE : short
    {
        DIK_0,
        DIK_1,
        DIK_2,
        DIK_3,
        DIK_4,
        DIK_5,
        DIK_6,
        DIK_7,
        DIK_8,
        DIK_9,
        DIK_A,
        DIK_ABNT_C1,
        DIK_ABNT_C2,
        DIK_ADD,
        DIK_APOSTROPHE,
        DIK_APPS,
        DIK_AT,
        DIK_AX,
        DIK_B,
        DIK_BACK,
        DIK_BACKSLASH,
        DIK_C,
        DIK_CALCULATOR,
        DIK_CAPITAL,
        DIK_COLON,
        DIK_COMMA,
        DIK_CONVERT,
        DIK_D,
        DIK_DECIMAL,
        DIK_DELETE,
        DIK_DIVIDE,
        DIK_DOWN,
        DIK_E,
        DIK_END,
        DIK_EQUALS,
        DIK_ESCAPE,
        DIK_F,
        DIK_F1,
        DIK_F2,
        DIK_F3,
        DIK_F4,
        DIK_F5,
        DIK_F6,
        DIK_F7,
        DIK_F8,
        DIK_F9,
        DIK_F10,
        DIK_F11,
        DIK_F12,
        DIK_F13,
        DIK_F14,
        DIK_F15,
        DIK_G,
        DIK_GRAVE,
        DIK_H,
        DIK_HOME,
        DIK_I,
        DIK_INSERT,
        DIK_J,
        DIK_K,
        DIK_KANA,
        DIK_KANJI,
        DIK_L,
        DIK_LBRACKET,
        DIK_LCONTROL,
        DIK_LEFT,
        DIK_LMENU,
        DIK_LSHIFT,
        DIK_LWIN,
        DIK_M,
        DIK_MAIL,
        DIK_MEDIASELECT,
        DIK_MEDIASTOP,
        DIK_MINUS,
        DIK_MULTIPLY,
        DIK_MUTE,
        DIK_MYCOMPUTER,
        DIK_N,
        DIK_NEXT,
        DIK_NEXTTRACK,
        DIK_NOCONVERT,
        DIK_NUMLOCK,
        DIK_NUMPAD0,
        DIK_NUMPAD1,
        DIK_NUMPAD2,
        DIK_NUMPAD3,
        DIK_NUMPAD4,
        DIK_NUMPAD5,
        DIK_NUMPAD6,
        DIK_NUMPAD7,
        DIK_NUMPAD8,
        DIK_NUMPAD9,
        DIK_NUMPADCOMMA,
        DIK_NUMPADENTER,
        DIK_NUMPADEQUALS,
        DIK_O,
        DIK_OEM_102,
        DIK_P,
        DIK_PAUSE,
        DIK_PERIOD,
        DIK_PLAYPAUSE,
        DIK_POWER,
        DIK_PREVTRACK,
        DIK_PRIOR,
        DIK_Q,
        DIK_R,
        DIK_RBRACKET,
        DIK_RCONTROL,
        DIK_RETURN,
        DIK_RIGHT,
        DIK_RMENU,
        DIK_RSHIFT,
        DIK_RWIN,
        DIK_S,
        DIK_SCROLL,
        DIK_SEMICOLON,
        DIK_SLASH,
        DIK_SLEEP,
        DIK_SPACE,
        DIK_STOP,
        DIK_SUBTRACT,
        DIK_SYSRQ,
        DIK_T,
        DIK_TAB,
        DIK_U,
        DIK_UNDERLINE,
        DIK_UNLABELED,
        DIK_UP,
        DIK_V,
        DIK_VOLUMEDOWN,
        DIK_VOLUMEUP,
        DIK_W,
        DIK_WAKE,
        DIK_WEBBACK,
        DIK_WEBFAVORITES,
        DIK_WEBFORWARD,
        DIK_WEBHOME,
        DIK_WEBREFRESH,
        DIK_WEBSEARCH,
        DIK_WEBSTOP,
        DIK_X,
        DIK_Y,
        DIK_YEN,
        DIK_Z,
    };

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

    //[StructLayout(LayoutKind.Explicit)]
    //private struct INPUT
    //{
    //    [FieldOffset(0)]
    //    public int type;
    //    [FieldOffset(4)]
    //    public MOUSEINPUT mi;
    //    [FieldOffset(4)]
    //    public KEYBDINPUT ki;
    //    [FieldOffset(4)]
    //    public HARDWAREINPUT hi;
    //}

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

    public class SendInputSharp
    {
        [DllImport("user32.dll")]
        private static extern uint SendInput(
            uint nInputs,                                               // 入力イベントの数 
            [MarshalAs(UnmanagedType.LPArray), In] INPUT[] pInputs,     // 挿入する入力イベントの配列
            int cbSize                                                  // 構造体のサイズ
        );

        [DllImport("user32.dll")]
        private extern static IntPtr GetMessageExtraInfo();

        [DllImport("user32.dll")]
        private extern static uint MapVirtualKey(
            uint uCode,     // 仮想キーコードまたはスキャンコード
            uint uMapType   // 実行したい変換の種類
        );

        private static readonly List<KeyValuePair<VIRTUALKEYCODE, SCANCODE>> KeyCodePairList = new List<KeyValuePair<VIRTUALKEYCODE,SCANCODE>>
        {
            new KeyValuePair<VIRTUALKEYCODE, SCANCODE>(VIRTUALKEYCODE.VK_UP, SCANCODE.DIK_UP), 
            new KeyValuePair<VIRTUALKEYCODE, SCANCODE>(VIRTUALKEYCODE.VK_DOWN, SCANCODE.DIK_DOWN), 
            new KeyValuePair<VIRTUALKEYCODE, SCANCODE>(VIRTUALKEYCODE.VK_LEFT, SCANCODE.DIK_LEFT), 
            new KeyValuePair<VIRTUALKEYCODE, SCANCODE>(VIRTUALKEYCODE.VK_RIGHT, SCANCODE.DIK_RIGHT), 
            new KeyValuePair<VIRTUALKEYCODE, SCANCODE>(VIRTUALKEYCODE.VK_INSERT, SCANCODE.DIK_INSERT), 
            new KeyValuePair<VIRTUALKEYCODE, SCANCODE>(VIRTUALKEYCODE.VK_HOME, SCANCODE.DIK_HOME), 
            new KeyValuePair<VIRTUALKEYCODE, SCANCODE>(VIRTUALKEYCODE.VK_PRIOR, SCANCODE.DIK_PRIOR),
            new KeyValuePair<VIRTUALKEYCODE, SCANCODE>(VIRTUALKEYCODE.VK_DELETE, SCANCODE.DIK_DELETE), 
            new KeyValuePair<VIRTUALKEYCODE, SCANCODE>(VIRTUALKEYCODE.VK_END, SCANCODE.DIK_END), 
            new KeyValuePair<VIRTUALKEYCODE, SCANCODE>(VIRTUALKEYCODE.VK_NEXT, SCANCODE.DIK_NEXT),
            new KeyValuePair<VIRTUALKEYCODE, SCANCODE>(VIRTUALKEYCODE.VK_LMENU, SCANCODE.DIK_LMENU), 
            new KeyValuePair<VIRTUALKEYCODE, SCANCODE>(VIRTUALKEYCODE.VK_RMENU, SCANCODE.DIK_RMENU), 
            new KeyValuePair<VIRTUALKEYCODE, SCANCODE>(VIRTUALKEYCODE.VK_LCONTROL, SCANCODE.DIK_LCONTROL), 
            new KeyValuePair<VIRTUALKEYCODE, SCANCODE>(VIRTUALKEYCODE.VK_RCONTROL, SCANCODE.DIK_RCONTROL),
            new KeyValuePair<VIRTUALKEYCODE, SCANCODE>(VIRTUALKEYCODE.VK_DIVIDE, SCANCODE.DIK_DIVIDE)
        };

        private static SCANCODE VirtualToScan(VIRTUALKEYCODE code) 
        {
            uint outKeyCode = MapVirtualKey((uint)code, 0);
            foreach (var c in KeyCodePairList.Where(p => p.Key == code).Select(p => p.Value))
            {
                return c;
            }
            return (SCANCODE)outKeyCode;
        }

        private static VIRTUALKEYCODE ScanToVirtual(SCANCODE code)
        {
            uint outKeyCode = MapVirtualKey((uint)code, 1);
            foreach (var c in KeyCodePairList.Where(p => p.Value == code).Select(p => p.Key))
            {
                return c;
            }
            return (VIRTUALKEYCODE)outKeyCode;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="keyEvents"></param>
        /// <returns></returns>
        public static uint SendKeybordInput(IEnumerable<KeyValuePair<VIRTUALKEYCODE, KEYEVENTFLAG>> keyEvents)
        {
            var inputs = keyEvents.Select((p) =>
            {
                return new INPUT()
                {
                    type = (int)INPUTTYPE.KEYBOARD,
                    U = new InputUnion(){ ki = new KEYBDINPUT() 
                    {
                        wVk = (ushort)p.Key,
                        wScan = (ushort)VirtualToScan(p.Key),
                        dwFlags = (int)p.Value,
                        time = 0,
                        dwExtraInfo = GetMessageExtraInfo()
                    }}
                };
            }).ToArray();
            return SendInput((uint)inputs.Length, inputs, INPUT.Size);
        }
    }
}
