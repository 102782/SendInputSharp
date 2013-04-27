using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace SendInputSharp
{
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

        [DllImport("user32.dll")]
        private static extern ushort GetAsyncKeyState(
            int nVirtKey    // 仮想キーコード
        );

        /// <summary>
        /// 仮想キーコードで指定したキーが押されているかどうかをbool値で返します。
        /// </summary>
        /// <param name="vKey">仮想キーコード</param>
        /// <returns></returns>
        public static bool IsKeyPushedDown(VIRTUALKEYCODE vKey)
        {
            return 0 != (GetAsyncKeyState((int)vKey) & 0x8000);
        }

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
        /// イベントを順にキーボード入力ストリームに挿入します。イベントは仮想キーコードと入力状態フラグのペアで指定します。
        /// </summary>
        /// <param name="keyEvents">
        /// 挿入するイベントのリスト
        /// </param>
        /// <returns>
        /// キーボード入力ストリームへ挿入することができたイベントの数を返します。ほかのスレッドによって入力がすでにブロックされている場合、関数は 0 を返します。拡張エラー情報を取得するには、GetLastError 関数を使います。
        /// </returns>
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
