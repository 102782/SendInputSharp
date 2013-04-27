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

        private static readonly List<KeyValuePair<Keys, SCANCODE>> KeyCodePairList = new List<KeyValuePair<Keys, SCANCODE>>
        {
            new KeyValuePair<Keys, SCANCODE>(Keys.Up, SCANCODE.DIK_UP), 
            new KeyValuePair<Keys, SCANCODE>(Keys.Down, SCANCODE.DIK_DOWN), 
            new KeyValuePair<Keys, SCANCODE>(Keys.Left, SCANCODE.DIK_LEFT), 
            new KeyValuePair<Keys, SCANCODE>(Keys.Right, SCANCODE.DIK_RIGHT), 
            new KeyValuePair<Keys, SCANCODE>(Keys.Insert, SCANCODE.DIK_INSERT), 
            new KeyValuePair<Keys, SCANCODE>(Keys.Home, SCANCODE.DIK_HOME), 
            new KeyValuePair<Keys, SCANCODE>(Keys.Prior, SCANCODE.DIK_PRIOR),
            new KeyValuePair<Keys, SCANCODE>(Keys.Delete, SCANCODE.DIK_DELETE), 
            new KeyValuePair<Keys, SCANCODE>(Keys.End, SCANCODE.DIK_END), 
            new KeyValuePair<Keys, SCANCODE>(Keys.Next, SCANCODE.DIK_NEXT),
            new KeyValuePair<Keys, SCANCODE>(Keys.LMenu, SCANCODE.DIK_LMENU), 
            new KeyValuePair<Keys, SCANCODE>(Keys.RMenu, SCANCODE.DIK_RMENU), 
            new KeyValuePair<Keys, SCANCODE>(Keys.LControlKey, SCANCODE.DIK_LCONTROL), 
            new KeyValuePair<Keys, SCANCODE>(Keys.RControlKey, SCANCODE.DIK_RCONTROL),
            new KeyValuePair<Keys, SCANCODE>(Keys.Divide, SCANCODE.DIK_DIVIDE)
        };

        private static SCANCODE VirtualToScan(Keys code) 
        {
            uint outKeyCode = MapVirtualKey((uint)code, 0);
            foreach (var c in KeyCodePairList.Where(p => p.Key == code).Select(p => p.Value))
            {
                return c;
            }
            return (SCANCODE)outKeyCode;
        }

        private static Keys ScanToVirtual(SCANCODE code)
        {
            uint outKeyCode = MapVirtualKey((uint)code, 1);
            foreach (var c in KeyCodePairList.Where(p => p.Value == code).Select(p => p.Key))
            {
                return c;
            }
            return (Keys)outKeyCode;
        }

        /// <summary>
        /// イベントを順にキーボード入力ストリームに挿入します。イベントは仮想キーコードと入力状態フラグのペアで指定します。
        /// </summary>
        /// <param name="keyEvents">
        /// 挿入するイベントのリスト
        /// </param>
        /// <returns>
        /// キーボード入力ストリームへ挿入することができたイベントの数を返します。ほかのスレッドによって入力がすでにブロックされている場合、関数は 0 を返します。拡張エラー情報を取得するには、GetLastError 関数を使います。
        /// </returns>
        public static uint SendKeybordInput(IEnumerable<KeyValuePair<Keys, KEYEVENTFLAG>> keyEvents)
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
