using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SendInputSharp
{
    public abstract class InputEvent
    {
        internal abstract INPUT ToInput();
    }

    public class KeyInputEvent : InputEvent
    {
        public Keys virtualKeyCode { get; set; }
        public KEYEVENTFLAG flag { get; set; }

        /// <summary>
        /// キーボード入力イベント
        /// </summary>
        /// <param name="virtualKeyCode">仮想キーコード</param>
        /// <param name="flag">入力イベントフラグ</param>
        public KeyInputEvent(Keys virtualKeyCode, KEYEVENTFLAG flag)
        {
            this.virtualKeyCode = virtualKeyCode;
            this.flag = flag;
        }

        internal override INPUT ToInput()
        {
            return new INPUT()
            {
                type = (int)INPUTTYPE.KEYBOARD,
                U = new InputUnion()
                {
                    ki = new KEYBDINPUT()
                    {
                        wVk = (ushort)this.virtualKeyCode,
                        wScan = (ushort)ManagedSendInput.VirtualToScan(this.virtualKeyCode),
                        dwFlags = (int)this.flag,
                        time = 0,
                        dwExtraInfo = ManagedSendInput.GetMessageExtraInfo()
                    }
                }
            };
        }
    }

    public class MouseInputEvent : InputEvent
    {
        public int dx { get; set; }
        public int dy { get; set; }
        public int wheel { get; set; }
        public MOUSEEVENTFLAG flag { get; set; }

        /// <summary>
        /// マウス入力イベント
        /// </summary>
        /// <param name="dx">カーソルのX座標を相対座標または絶対座標で指定します。</param>
        /// <param name="dy">カーソルのY座標を相対座標または絶対座標で指定します。</param>
        /// <param name="wheel">ホイール回転量</param>
        /// <param name="flag">入力イベントフラグ</param>
        public MouseInputEvent(int dx, int dy, int wheel, MOUSEEVENTFLAG flag)
        {
            this.dx = dx;
            this.dy = dy;
            this.wheel = wheel;
            this.flag = flag;
        }

        internal override INPUT ToInput()
        {
            return new INPUT()
            {
                type = (int)INPUTTYPE.MOUSE,
                U = new InputUnion()
                {
                    mi = new MOUSEINPUT()
                    {
                        dx = this.dx,
                        dy = this.dy,
                        mouseData = this.wheel,
                        dwFlags = (int)this.flag,
                        time = 0,
                        dwExtraInfo = ManagedSendInput.GetMessageExtraInfo()
                    }
                }
            };
        }
    }

    public class ManagedSendInput
    {
        #region DllImports
        [DllImport("user32.dll")]
        private static extern uint SendInput(
            uint nInputs,                                               // 入力イベントの数 
            [MarshalAs(UnmanagedType.LPArray), In] INPUT[] pInputs,     // 挿入する入力イベントの配列
            int cbSize                                                  // 構造体のサイズ
        );

        [DllImport("user32.dll")]
        internal extern static IntPtr GetMessageExtraInfo();

        [DllImport("user32.dll")]
        private extern static uint MapVirtualKey(
            uint uCode,     // 仮想キーコードまたはスキャンコード
            uint uMapType   // 実行したい変換の種類
        );

        [DllImport("user32.dll")]
        private static extern ushort GetAsyncKeyState(
            int nVirtKey    // 仮想キーコード
        );
        #endregion

        #region KeyCodePairList
        private static readonly Tuple<Keys, SCANCODE>[] KeyCodePairList = new []
        {
            new Tuple<Keys, SCANCODE>(Keys.Up, SCANCODE.Up), 
            new Tuple<Keys, SCANCODE>(Keys.Down, SCANCODE.Down), 
            new Tuple<Keys, SCANCODE>(Keys.Left, SCANCODE.Left), 
            new Tuple<Keys, SCANCODE>(Keys.Right, SCANCODE.Right), 
            new Tuple<Keys, SCANCODE>(Keys.Insert, SCANCODE.Insert), 
            new Tuple<Keys, SCANCODE>(Keys.Home, SCANCODE.Home), 
            new Tuple<Keys, SCANCODE>(Keys.Prior, SCANCODE.Prior),
            new Tuple<Keys, SCANCODE>(Keys.Delete, SCANCODE.Delete), 
            new Tuple<Keys, SCANCODE>(Keys.End, SCANCODE.End), 
            new Tuple<Keys, SCANCODE>(Keys.Next, SCANCODE.Next),
            new Tuple<Keys, SCANCODE>(Keys.LMenu, SCANCODE.LeftMenu), 
            new Tuple<Keys, SCANCODE>(Keys.RMenu, SCANCODE.RightMenu), 
            new Tuple<Keys, SCANCODE>(Keys.LControlKey, SCANCODE.LeftControl), 
            new Tuple<Keys, SCANCODE>(Keys.RControlKey, SCANCODE.RightControl),
            new Tuple<Keys, SCANCODE>(Keys.Divide, SCANCODE.Divide)
        };
        #endregion

        internal static SCANCODE VirtualToScan(Keys code) 
        {
            uint outKeyCode = MapVirtualKey((uint)code, 0);
            foreach (var c in KeyCodePairList.Where(p => p.Item1 == code).Select(p => p.Item2))
            {
                return c;
            }
            return (SCANCODE)outKeyCode;
        }

        internal static Keys ScanToVirtual(SCANCODE code)
        {
            uint outKeyCode = MapVirtualKey((uint)code, 1);
            foreach (var c in KeyCodePairList.Where(p => p.Item2 == code).Select(p => p.Item1))
            {
                return c;
            }
            return (Keys)outKeyCode;
        }

        /// <summary>
        /// イベントを入力ストリームに挿入します。
        /// </summary>
        /// <param name="events">挿入するイベントのリスト</param>
        /// <returns>
        /// 入力ストリームへ挿入することができたイベントの数を返します。ほかのスレッドによって入力がすでにブロックされている場合、関数は 0 を返します。拡張エラー情報を取得するには、GetLastError 関数を使います。
        /// </returns>
        public static uint SendInput(IEnumerable<InputEvent> events)
        {
            var inputs = events.Select(p => p.ToInput()).ToArray();
            return SendInput((uint)inputs.Length, inputs, INPUT.Size);
        }

        /// <summary>
        /// イベントをキーボード入力ストリームに挿入します。
        /// </summary>
        /// <param name="virtualKey">仮想キーコード</param>
        /// <param name="flag">入力イベントフラグ</param>
        /// <returns>
        /// キーボード入力ストリームへ挿入することができたイベントの数を返します。ほかのスレッドによって入力がすでにブロックされている場合、関数は 0 を返します。拡張エラー情報を取得するには、GetLastError 関数を使います。
        /// </returns>
        public static uint SendKeybordInput(Keys virtualKey, KEYEVENTFLAG eventFlag)
        {
            return SendInput(new[] { new KeyInputEvent(virtualKey, eventFlag) });
        }

        /// <summary>
        /// イベントをマウス入力ストリームに挿入します。
        /// </summary>
        /// <param name="dx">カーソルのX座標を相対座標または絶対座標で指定します。</param>
        /// <param name="dy">カーソルのY座標を相対座標または絶対座標で指定します。</param>
        /// <param name="wheel">ホイール回転量</param>
        /// <param name="flag">入力イベントフラグ</param>
        /// <returns>
        /// マウス入力ストリームへ挿入することができたイベントの数を返します。ほかのスレッドによって入力がすでにブロックされている場合、関数は 0 を返します。拡張エラー情報を取得するには、GetLastError 関数を使います。
        /// </returns>
        public static uint SendMouseInput(int dx, int dy, int wheel, MOUSEEVENTFLAG flag)
        {
            return SendInput(new[] { new MouseInputEvent(dx, dy, wheel, flag) });
        }

        /// <summary>
        /// 呼び出し側のスレッドが持つ最新のエラーコードを取得します。
        /// </summary>
        /// <returns></returns>
        public static int GetLastError()
        {
            return Marshal.GetLastWin32Error();
        }

        /// <summary>
        /// 仮想キーコードで指定したキーが押されているかどうかをbool値で返します。
        /// </summary>
        /// <param name="vKey">仮想キーコード</param>
        /// <returns></returns>
        public static bool IsKeyPushedDown(Keys vKey)
        {
            return 0 != (GetAsyncKeyState((int)vKey) & 0x8000);
        }
    }
}
