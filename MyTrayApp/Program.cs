using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Text;
using System.IO.Ports;
using System.Threading;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace MyTrayApp
{
    public class SysTrayApp : Form
    {
        static SerialPort serialPort = null;
        
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("kernel32.dll")]
        static extern IntPtr LoadLibrary(string lpFileName);

        // Get a handle to an application window.
        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        // Activate an application window.
        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        
        [STAThread]
        public static void Main()
        {
            Application.Run(new SysTrayApp());
        }

        private NotifyIcon trayIcon;
        private ContextMenu trayMenu;

        public SysTrayApp()
        {
            // Create a simple tray menu with only one item.
            trayMenu = new ContextMenu();                        
            trayMenu.MenuItems.Add("Стоп", OnStop);
            trayMenu.MenuItems.Add("Тест данных", OnTest);
            trayMenu.MenuItems.Add("Старт", OnStart);
            trayMenu.MenuItems.Add("Выход", OnExit);

            // Create a tray icon. In this example we use a
            // standard system icon for simplicity, but you
            // can of course use your own custom icon too.
            trayIcon = new NotifyIcon();
            trayIcon.Text = "Вес из весов";
            trayIcon.Icon = new Icon(SystemIcons.Application, 40, 40);

            // Add menu to tray icon and show it.
            trayIcon.ContextMenu = trayMenu;
            trayIcon.Visible = true;
            hhook = SetHook(_proc);
        }

        private void OnTest(object sender, EventArgs e) 
        {
            if (!serialPort.IsOpen) 
            {
                MessageBox.Show("Ошибка открытия COM1", "RS read weight", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }            

            string result = GetStringWeight();
            result = DoRight(result);                            
            MessageBox.Show(result, "RS read weight", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private bool OpenCommPort() 
        {
            if (serialPort != null)
                if (serialPort.IsOpen)
                    return true;

            serialPort = new SerialPort("COM1", 9600);
            serialPort.Parity = Parity.None;
            serialPort.DataBits = 8;
            serialPort.DtrEnable = true;
            serialPort.RtsEnable = true;
            serialPort.Handshake = Handshake.XOnXOff;
            serialPort.Open();

            return serialPort.IsOpen;
        }

        private void CloseCommPort() 
        {
            if (serialPort != null)
                if (serialPort.IsOpen)
                    serialPort.Close();
        }

        static private string GetStringWeight() 
        {
            byte[] request = new byte[] { 0x56, 0x45, 0x53, 0x0D };

            serialPort.Write(request, 0, 4);

            Thread.Sleep(500);

            char[] buf = new char[8];
            serialPort.Read(buf, 0, 8);

            string result = new string(buf);

            result = Regex.Replace(result, @"\s+", "");

            return result;
        }

        static private string DoRight(string value) 
        {
            double val = 0;
            try
            {
                val = Int32.Parse(value);
            }
            catch(FormatException ex)
            {
                MessageBox.Show(value, "RS read weight", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "Error";
            }
            double tens = 10 * ((int)(val / 10));
            double ost = val - tens;
            if (ost > 5) tens += 10;
            return tens.ToString();
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        const int WH_KEYBOARD_LL = 13; // Номер глобального LowLevel-хука на клавиатуру
        const int WM_KEYDOWN = 0x100; // Сообщения нажатия клавиши

        private LowLevelKeyboardProc _proc = hookProc;

        private static IntPtr hhook = IntPtr.Zero;

        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
                    GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        //public void SetHook()
        //{
        //    IntPtr hInstance = LoadLibrary("User32");
        //    hhook = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, hInstance, 0);
        //}

        public static void UnHook()
        {
            UnhookWindowsHookEx(hhook);
        }

        public static IntPtr hookProc(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);

                //MessageBox.Show(vkCode.ToString());
                ////////ОБРАБОТКА НАЖАТИЯ F11
                if (vkCode.ToString() == "122")
                {
                    
                    string result = GetStringWeight();
                    result = DoRight(result);
                    //MessageBox.Show(result);
                    SendKeys.Send(result);
                    //if (!result.Equals("Error") && result != "Error")
                    //{
                    //    byte[] resbyte = Encoding.ASCII.GetBytes(result);
                    //    for (int i = 0; i < resbyte.Length; i++)
                    //    {
                    //        PressKey(resbyte[i], false);
                    //        PressKey(resbyte[i], true);
                    //    }
                    //}
                }
                return CallNextHookEx(hhook, code, wParam, lParam);
            }
            else
                return CallNextHookEx(hhook, code, wParam, lParam);
        }

        protected override void OnLoad(EventArgs e)
        {
            Visible = false; // Hide form window.
            ShowInTaskbar = false; // Remove from taskbar.            
            base.OnLoad(e);            
        }

        protected void initPort() {
         
        }

        static protected string getWeight() 
        {
            serialPort.WriteLine("VES");
            Thread.Sleep(1500);
            String result = serialPort.ReadLine();

            return result;
        }

        private void OnStart(object sender, EventArgs e) 
        {
           if (serialPort == null)
                if (!OpenCommPort())
                {
                    MessageBox.Show("Ошибка открытия COM1", "RS read weight", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else 
                {
                    MessageBox.Show("COM1 открыт", "RS read weight", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;                
                }
            //SetHook();
        }

        private void OnStop(object sender, EventArgs e)
        {
            //UnHook();
        }

        private void OnExit(object sender, EventArgs e)
        {
            UnHook();
            CloseCommPort();
            Application.Exit();
        }

        protected override void Dispose(bool isDisposing)
        {
            if (isDisposing)
            {
                // Release the icon resource.
                trayIcon.Dispose();
            }
            UnHook();
            CloseCommPort();
            base.Dispose(isDisposing);
        }
    }
}