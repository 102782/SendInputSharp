using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SendInputSharp
{
    public class ManagedSendInput
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

        [DllImport("user32.dll")]
        private static extern ushort GetAsyncKeyState(
            int nVirtKey    // 仮想キーコード
        );

        /// <summary>
        /// 仮想キーコードで指定したキーが押されているかどうかをbool値で返します。
        /// </summary>
        /// <param name="vKey">仮想キーコード</param>
        /// <returns></returns>
        public static bool IsKeyPushedDown(Keys vKey)
        {
            return 0 != (GetAsyncKeyState((int)vKey) & 0x8000);
        }

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

        private static SCANCODE VirtualToScan(Keys code) 
        {
            uint outKeyCode = MapVirtualKey((uint)code, 0);
            foreach (var c in KeyCodePairList.Where(p => p.Item1 == code).Select(p => p.Item2))
            {
                return c;
            }
            return (SCANCODE)outKeyCode;
        }

        private static Keys ScanToVirtual(SCANCODE code)
        {
            uint outKeyCode = MapVirtualKey((uint)code, 1);
            foreach (var c in KeyCodePairList.Where(p => p.Item2 == code).Select(p => p.Item1))
            {
                return c;
            }
            return (Keys)outKeyCode;
        }

        /// <summary>
        /// イベントを順にキーボード入力ストリームに挿入します。イベントは仮想キーコードと入力イベントフラグのペアで指定します。
        /// </summary>
        /// <param name="keyEvents">
        /// 挿入するイベントのリスト
        /// </param>
        /// <returns>
        /// キーボード入力ストリームへ挿入することができたイベントの数を返します。ほかのスレッドによって入力がすでにブロックされている場合、関数は 0 を返します。拡張エラー情報を取得するには、GetLastError 関数を使います。
        /// </returns>
        public static uint SendKeybordInput(IEnumerable<Tuple<Keys, KEYEVENTFLAG>> keyEvents)
        {
            var inputs = keyEvents.Select((p) =>
            {
                return new INPUT()
                {
                    type = (int)INPUTTYPE.KEYBOARD,
                    U = new InputUnion(){ ki = new KEYBDINPUT() 
                    {
                        wVk = (ushort)p.Item1,
                        wScan = (ushort)VirtualToScan(p.Item1),
                        dwFlags = (int)p.Item2,
                        time = 0,
                        dwExtraInfo = GetMessageExtraInfo()
                    }}
                };
            }).ToArray();
            return SendInput((uint)inputs.Length, inputs, INPUT.Size);
        }

        /// <summary>
        /// イベントをキーボード入力ストリームに挿入します。イベントは仮想キーコードと入力イベントフラグのペアで指定します。
        /// </summary>
        /// <param name="virtualKey">仮想キーコード</param>
        /// <param name="flag">入力イベントフラグ</param>
        /// <returns>キーボード入力ストリームへ挿入することができたイベントの数を返します。ほかのスレッドによって入力がすでにブロックされている場合、関数は 0 を返します。拡張エラー情報を取得するには、GetLastError 関数を使います。</returns>
        public static uint SendKeybordInput(Keys virtualKey, KEYEVENTFLAG eventFlag)
        {
            return SendKeybordInput(new[] { new Tuple<Keys, KEYEVENTFLAG>(virtualKey, eventFlag) });
        }

        /// <summary>
        /// 呼び出し側のスレッドが持つ最新のエラーコードを取得します。
        /// </summary>
        /// <returns></returns>
        public static int GetLastError()
        {
            return Marshal.GetLastWin32Error();
        }
    }
}
