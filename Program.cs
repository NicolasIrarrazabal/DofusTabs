using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace DofusMiniTabber
{
    static class Program
    {
        private const string MutexName = "DofusMiniTabber_SingleInstance_v1";

        // Mensaje único global para "restaura tu ventana"
        internal static readonly uint WM_BRING_TO_FRONT =
            RegisterWindowMessage("DofusMiniTabber_BringToFront");

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern uint RegisterWindowMessage(string lpString);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindow(string? lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            using var mutex = new Mutex(true, MutexName, out bool createdNew);

            if (!createdNew)
            {
                // Ya existe una instancia → mandarle señal y salir silenciosamente
                IntPtr hWnd = FindWindow(null, "Wintabber Dofus");
                if (hWnd != IntPtr.Zero)
                    PostMessage(hWnd, WM_BRING_TO_FRONT, IntPtr.Zero, IntPtr.Zero);
                return;
            }

            Application.Run(new Form1());
        }
    }
}
