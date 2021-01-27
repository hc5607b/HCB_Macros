using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO;
using WindowsInput;
using WindowsInput.Native;

namespace HcbMacro
{
    public class LowLevelKeyboardHook
    {
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_KEYUP = 0x101;
        private const int WM_SYSKEYUP = 0x105;

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        public delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        public event EventHandler<Keys> OnKeyPressed;
        public event EventHandler<Keys> OnKeyUnpressed;

        private LowLevelKeyboardProc _proc;
        private IntPtr _hookID = IntPtr.Zero;

        public LowLevelKeyboardHook()
        {
            _proc = HookCallback;
        }

        public void HookKeyboard()
        {
            _hookID = SetHook(_proc);
        }

        public void UnHookKeyboard()
        {
            UnhookWindowsHookEx(_hookID);
        }

        private IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);

                OnKeyPressed.Invoke(this, ((Keys)vkCode));
            }
            else if (nCode >= 0 && wParam == (IntPtr)WM_KEYUP || wParam == (IntPtr)WM_SYSKEYUP)
            {
                int vkCode = Marshal.ReadInt32(lParam);

                OnKeyUnpressed.Invoke(this, ((Keys)vkCode));
            }

            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }
    }

    public class Program
    {
        static void Main(string[] args)
        {
            new App();
            Console.ReadLine();
        }
    }

    class App {
        List<Macro> macros = new List<Macro>();

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        bool visible = true;
        void toggleVisibility()
        {
            var handle = GetConsoleWindow();
            if (visible)
            {
                ShowWindow(handle, SW_HIDE);
                visible = false;
            }
            else
            {
                ShowWindow(handle, SW_SHOW);
                visible = true;
            }
        }

        void loadMacros() {
            List<FileInfo> files = new DirectoryInfo(Directory.GetCurrentDirectory() + "/macros").GetFiles().ToList();

            int loadCount = 0;
            foreach (FileInfo f in files)
            {
                if (f.Extension == ".mac") {
                    string name = f.Name.Split('.')[0];
                    string[] data = File.ReadAllLines(f.FullName);

                    List<Keys> comboList = new List<Keys>();
                    string combo = data[0];

                    foreach (string s in combo.Split('+')) {
                        Keys key;
                        Enum.TryParse(s, out key);
                        comboList.Add(key);
                    }

                    List<string> act = new List<string>();
                    for (int i = 1; i < data.Length; i++)
                    {
                        act.Add(data[i]);
                    }
                    
                    macros.Add(new Macro(name, comboList, act));
                    loadCount++;
                }
            }

            Console.WriteLine($"{loadCount} macros loaded");
            
        }

        bool listenForCombination = true;
        public App()
        {
            toggleVisibility();
            loadMacros();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            LowLevelKeyboardHook kbh = new LowLevelKeyboardHook();
            kbh.OnKeyPressed += kbh_OnKeyPressed;
            kbh.OnKeyUnpressed += kbh_OnKeyUnpressed;
            kbh.HookKeyboard();

            Application.Run();

            kbh.UnHookKeyboard();
            Console.WriteLine("Vituiks meni");
        }
        List<Keys> pressedKeys = new List<Keys>();
        void kbh_OnKeyPressed(object sender, Keys e)
        {
            pressedKeys.Add(e);
        }
        void kbh_OnKeyUnpressed(object sender, Keys e)
        {
            if (listenForCombination) { chkCombo(); }
            //pressedKeys.Remove(e);
            pressedKeys.Clear();
            Console.WriteLine(e + " " + pressedKeys.Count);
        }

        void chkCombo() {
            Macro m = getMacroByCombo(pressedKeys);
            if (m != null) { Console.WriteLine(m.ToString()); Task.Run(() => executeMacro(m)); }
        }

        Macro getMacroByCombo(List<Keys> curMacro)
        {
            if (string.Join(", ", curMacro) == "LControlKey, LShiftKey, F12") { toggleVisibility(); return null; }
            foreach (Macro m in macros)
            {
                if (string.Join(", ", curMacro) == string.Join(", ", m.GetCombo())) { return m; }
            }
            return null;
        }

        async Task executeMacro(Macro m)
        {
            Thread.Sleep(500);
            foreach (string s in m.GetActions())
            {
                InputSimulator ins = new InputSimulator();
                string type = s.Split('=')[0];
                string data = s.Substring(s.IndexOf('=') + 1);

                switch (type)
                {
                    case "COM":
                        List<VirtualKeyCode> keys = new List<VirtualKeyCode>();
                        foreach (string keycode in data.Split('+'))
                        {
                            VirtualKeyCode key;
                            Enum.TryParse(keycode, out key);
                            keys.Add(key);
                        }

                        foreach (VirtualKeyCode vkc in keys)
                        {
                            ins.Keyboard.KeyDown(vkc);
                            Thread.Sleep(10);
                        }
                        foreach (VirtualKeyCode vkc in keys)
                        {
                            ins.Keyboard.KeyUp(vkc);
                        }

                        break;

                    case "TXT":
                        ins.Keyboard.TextEntry(data);
                        break;

                    case "KEY":
                        VirtualKeyCode key1;
                        Enum.TryParse(data, out key1);
                        ins.Keyboard.KeyPress(key1);
                        break;

                    case "WAIT":
                        int waitTime = 0;
                        int.TryParse(data, out waitTime);
                        Thread.Sleep(waitTime);
                        break;
                }
                Thread.Sleep(10);
            }
        }
    }

    class Macro {
        string name;
        List<Keys> combo = new List<Keys>();
        List<string> actions = new List<string>();

        public List<Keys> GetCombo() { return combo; }
        public List<String> GetActions() { return actions; }
        public override string ToString()
        {
            return name;
        }

        public Macro(string _name, List<Keys> _comb, List<string> _actions) { combo = _comb;name = _name; actions = _actions; }
    }
}
