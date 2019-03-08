using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WatiN.Core;


namespace ConsoleApp.Helper
{
    public static class IEBrowserHelper
    {
        //private System.Configuration.AppSettingsReader appReader = new System.Configuration.AppSettingsReader();

        private static readonly int openedBrowsersLimit = 5;
        private static readonly int browserTimeout = 300;

        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);


        [MethodImpl(MethodImplOptions.Synchronized)]
        public static WatiN.Core.IE GetBrowser()
        {
            Settings.AttachToBrowserTimeOut = 240;
            Settings.WaitUntilExistsTimeOut = 240;
            Settings.WaitForCompleteTimeOut = 240;
            WatiN.Core.IE browser = null;
 

            VerifyDisponibility();
            try
            {//Primeira tentativa de criar o browser
                browser = CreateIEBrowser();
            }
            catch
            {//Caso dê algum problema tenta criar o browser mais uma vez
                System.Threading.Thread.Sleep(new TimeSpan(0, 0, 5));
                VerifyDisponibility();
                browser = CreateIEBrowser();
            }

            return browser;
        }

        public static void TypeTextForced(this TextField textField, string text)
        {

            SetForegroundWindow(textField.DomContainer.hWnd);
            SetFocus(textField.DomContainer.hWnd);
            textField.Focus();
            System.Windows.Forms.SendKeys.SendWait(text);
            Thread.Sleep(1000);
        }

        public static void PressEnterForced(this TextField textField)
        {
            SetForegroundWindow(textField.DomContainer.hWnd);
            SetFocus(textField.DomContainer.hWnd);
            textField.Focus();
            System.Windows.Forms.SendKeys.SendWait("{ENTER}");
            Thread.Sleep(1000);

        }
        public static void PressEscForced(this TextField textField)
        {
            SetForegroundWindow(textField.DomContainer.hWnd);
            SetFocus(textField.DomContainer.hWnd);
            textField.Focus();
            System.Windows.Forms.SendKeys.SendWait("{ESC}");
            Thread.Sleep(1000);

        }
        public static void PressTabcForced(this TextField textField)
        {
            SetForegroundWindow(textField.DomContainer.hWnd);
            SetFocus(textField.DomContainer.hWnd);
            textField.Focus();
            System.Windows.Forms.SendKeys.SendWait("{TAB}");
            Thread.Sleep(1000);

        }
        private static void VerifyDisponibility()
        {
            KillIEProcesses();
            int count = 1;
            while (Process.GetProcessesByName("IExplore").Count() > openedBrowsersLimit)
            {
                if (count % 100 == 0)//Em uma determinada quantidade de interações sem sucesso tenta matar os procesos para que não haja deadlock
                {
                    KillIEProcesses();
                }
            }

        }

        private static Process CreateIExploreInNewProcess()
        {
            var arguments = "about:blank";

            arguments = "-noframemerging " + arguments;

            Process m_Proc = null;
            try
            {
                m_Proc = Process.Start("IExplore.exe", arguments + "\n");
                if (m_Proc == null)
                {
                    throw new WatiN.Core.Exceptions.WatiNException("Could not start IExplore.exe process");
                }

            }
            catch (Exception e)
            {
                throw e;
            }

            return m_Proc;
        }

        private static void KillIEProcesses()
        {
            foreach (var p in Process.GetProcessesByName("IExplore"))
            {
                TimeSpan duration = DateTime.Now.Subtract(p.StartTime);
                int processTotalDuration = duration.Minutes * 60 + duration.Seconds;
                if (processTotalDuration > browserTimeout)
                {
                    p.Kill();
                }
            }
        }

        class IeWindowFinder
        {
            #region Interop
            [DllImport("user32.dll", SetLastError = true)]
            static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
            public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
            [DllImport("user32", EntryPoint = "GetClassNameA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
            internal static extern int GetClassName(IntPtr handleToWindow, StringBuilder className, int maxClassNameLength);
            #endregion

            readonly Process IeProcess;
            IntPtr HWnd = IntPtr.Zero;

            public IeWindowFinder(Process ieProcess)
            {
                this.IeProcess = ieProcess;
            }

            public IntPtr Find()
            {
                EnumWindows(FindIeWindowCallback, IntPtr.Zero);
                return HWnd;
            }

            bool FindIeWindowCallback(IntPtr hWnd, IntPtr lParam)
            {
                uint processId;
                GetWindowThreadProcessId(hWnd, out processId);

                if (processId == IeProcess.Id)
                {
                    int maxCapacity = 255;
                    var sbClassName = new StringBuilder(maxCapacity);
                    var lRes = GetClassName(hWnd, sbClassName, maxCapacity);
                    string className = lRes == 0 ? String.Empty : sbClassName.ToString();

                    if (className == "IEFrame")
                    {
                        this.HWnd = hWnd;
                        return false;
                    }
                }
                return true;
            }
        }

        private static WatiN.Core.IE CreateIEBrowser()
        { 
            Process ieProcess = CreateIExploreInNewProcess();

            IeWindowFinder findWindow = new IeWindowFinder(ieProcess);

            var action = new WatiN.Core.UtilityClasses.TryFuncUntilTimeOut(TimeSpan.FromSeconds(WatiN.Core.Settings.AttachToBrowserTimeOut))
            {
                SleepTime = TimeSpan.FromMilliseconds(500)
            };

            IntPtr hWnd = action.Try(() =>
            {
                return findWindow.Find();
            });

            ieProcess.Refresh();

            WatiN.Core.IE ie = WatiN.Core.IE.AttachTo<WatiN.Core.IE>(
                new WatiN.Core.Constraints.AttributeConstraint("hwnd", hWnd.ToString()), 5);

            return ie;
        }
 
    }
}
