using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelVBA
{
    class KeySendPassword
    {
        Login password = new Login();

        public string swPassword;

        [DllImport("user32.dll", EntryPoint = "FindWindow")]
        private static extern IntPtr FindWindow(string lp1, string lp2);

        [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool BringWindowToTop(IntPtr hWnd);
        public void Klucze()
        {
            password.ShowDialog();
            swPassword = password.pas;
            IntPtr hWnd1 = FindWindow("XLMAIN", null);
            if (hWnd1 != IntPtr.Zero)
            {
                bool ret = SetForegroundWindow(hWnd1); //Bring VBE to top.
            }

            SendKeys.SendWait("%{F11}");
            SendKeys.SendWait("^r");
           

            IntPtr hWnd = FindWindow("MsoCommandBar", null);
            if (hWnd != IntPtr.Zero)
            {
                bool ret = SetForegroundWindow(hWnd); //Bring VBE to top.
            }
            // string swPassword = "asdzxc";
            

            SendKeys.SendWait("{PGDN}");
            SendKeys.SendWait("{ENTER}");
            
            foreach (char c in swPassword)
                SendKeys.SendWait(c.ToString());            
            SendKeys.SendWait("{ENTER}");
            SendKeys.SendWait("{PGUP}");
            SendKeys.SendWait("{ENTER}");

            foreach (char c in swPassword)
                SendKeys.SendWait(c.ToString());
            SendKeys.SendWait("{ENTER}");
        }
        public void KluczS(ref Microsoft.Office.Interop.Excel.Application app)
        {
            password.ShowDialog();
            swPassword = password.pas;
            IntPtr hWnd1 = FindWindow("XLMAIN", null);
            if (hWnd1 != IntPtr.Zero)
            {
                bool ret = SetForegroundWindow(hWnd1); //Bring VBE to top.
            }
            app.ScreenUpdating = false;
            SendKeys.SendWait("%{F11}");
            SendKeys.SendWait("^r");
           
            IntPtr hWnd = FindWindow("MsoCommandBar", null);
            if (hWnd != IntPtr.Zero)
            {
                bool ret = SetForegroundWindow(hWnd); //Bring VBE to top.
            }
            // string swPassword = "asdzxc";
            hWnd = FindWindow("MsoCommandBar", null);
            SendKeys.SendWait("{PGDN}");
                
            SendKeys.SendWait("{ENTER}");
            
            foreach (char c in swPassword)
               SendKeys.SendWait(c.ToString());
            SendKeys.SendWait("{ENTER}");
            SendKeys.SendWait("{PGDN}");
            SendKeys.SendWait("{ENTER}");
            FindWindow("MsoCommandBar", null);
            if (hWnd != IntPtr.Zero)
            {
                bool ret = SetForegroundWindow(hWnd); //Bring VBE to top.
            }


        }
        public void KluczD()
        {
            password.ShowDialog();
            swPassword = password.pas;
            IntPtr hWnd1 = FindWindow("XLMAIN", null);
            if (hWnd1 != IntPtr.Zero)
            {
                bool ret = SetForegroundWindow(hWnd1); //Bring VBE to top.
            }

            SendKeys.SendWait("%{F11}");
            SendKeys.SendWait("^r");
            
            IntPtr hWnd = FindWindow("MsoCommandBar", null);
            if (hWnd != IntPtr.Zero)
            {
                bool ret = SetForegroundWindow(hWnd); //Bring VBE to top.
            }
            // string swPassword = "asdzxc";
            hWnd = FindWindow("MsoCommandBar", null);

            SendKeys.SendWait("{PGUP}");
            SendKeys.SendWait("{ENTER}");
            
            foreach (char c in swPassword)
              SendKeys.SendWait(c.ToString());
            SendKeys.SendWait("{ENTER}");

        }
    }
}
