using Microsoft.VisualBasic.Devices;
using System.ComponentModel;
using System.Diagnostics;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace Clicker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.TopMost = true;
        }

        // Import the necessary functions from user32.dll
        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetCursorPos(out POINT point);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);


        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public int mouseData;
            public int flags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        // Custom class to represent mouse point information
        public class MousePointInfo
        {
            public int MouseX { get; set; }
            public int MouseY { get; set; }
            public DateTime Time { get; set; }
            public double Duration { get; set; }
            public string TypeofAction { get; set; }
    }

        const uint MOUSEEVENTF_ABSOLUTE = 0x8000;
        const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        const uint MOUSEEVENTF_LEFTUP = 0x0004;
        const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;
        const uint MOUSEEVENTF_MIDDLEUP = 0x0040;
        const uint MOUSEEVENTF_MOVE = 0x0001;
        const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        const uint MOUSEEVENTF_RIGHTUP = 0x0010;
        const uint MOUSEEVENTF_XDOWN = 0x0080;
        const uint MOUSEEVENTF_XUP = 0x0100;
        const uint MOUSEEVENTF_WHEEL = 0x0800;
        const uint MOUSEEVENTF_HWHEEL = 0x01000;


        private const int WH_MOUSE_LL = 14;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_LBUTTONUP = 0x0202;
        private static IntPtr hookId = IntPtr.Zero;
        private static LowLevelMouseProc llMouseProc = HookCallback;

        // Declare globalMouseX at the class level
        private static int globalMouseX;
        private static int globalMouseY;

        // List to store information for each point
        private static List<MousePointInfo> mousePoints = new List<MousePointInfo>();

        private static int iStep =0;
        private static int iStepLearn = 0;

        private static bool bInLearning = false;

        private DateTime targetDate;

        //private static object lockmousePoints = new object();

        SemaphoreSlim semaphore = new(3);

        public void clickSimulator(Point p)
        {
            SetCursorPos(mousePoints[iStep].MouseX, mousePoints[iStep].MouseY);

            /*if ((int)mousePoints[iStep].Duration != 0)
            {
                timer1.Interval = (int)mousePoints[iStep].Duration * 1000;
            }*/

            label2.Text = timer1.Interval.ToString();
       
            if (mousePoints[iStep].TypeofAction == "DOWN")
            {
                mouse_event((int)(MOUSEEVENTF_LEFTDOWN), p.X, p.Y, 0, 0);   
            }
            if (mousePoints[iStep].TypeofAction == "UP")
            {  
                mouse_event((int)(MOUSEEVENTF_LEFTUP), p.X, p.Y, 0, 0);
            }
            label1.Text = p.ToString();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (mousePoints[iStep] is not null && (DateTime.Now >= targetDate))
            {
                targetDate = DateTime.Now.AddSeconds(mousePoints[iStep].Duration);
                clickSimulator(new Point(MousePosition.X, MousePosition.Y));
                iStep = iStep + 1;
            }
            else
            {
                iStep = 0;
                
            }
        }


        bool stop = true;
        private void button1_Click(object sender, EventArgs e)
        {

            stop = (stop) ? false : true;
            
            if (!stop)
            {
                timer1.Start();
                button1.BackColor = Color.LightGreen;
            }
            else
            {
                timer1.Stop();
                button1.BackColor = Color.LightGray;
                iStep = 0;
            }
        
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            // Set the global mouse hook
            hookId = SetHook(llMouseProc);
            Task.Run(() => HMI_Refresh());

            timer1.Interval = 200;
            //timer2.Interval = 1000;
        }

        private void HMI_Refresh()
        {
            //richTextBox1.Text = "";
            while (true)
            {
                Invoke(new Action(() => richTextBox1.Clear()));

                    foreach (MousePointInfo NPI in mousePoints)
                    {
                        Invoke(new Action(() => richTextBox1.AppendText(NPI.MouseX.ToString() + " " + NPI.MouseY.ToString() + " D: " + NPI.Duration.ToString() + " A: " + NPI.TypeofAction.ToString() + "\r\n")));
                        //richTextBox1.Text = richTextBox1.Text + NPI.MouseX.ToString() + " " + NPI.MouseY.ToString() + " D: " + NPI.Duration.ToString() + " A: " + NPI.TypeofAction.ToString() + "\r\n";
                    }
           
                Task.Delay(500).Wait();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            // Unhook the global mouse hook
            UnhookWindowsHookEx(hookId);
        }



        private static IntPtr SetHook(LowLevelMouseProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
         private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            MSLLHOOKSTRUCT hookStruct = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
            // Get the global mouse position on click
            globalMouseX = hookStruct.pt.x;
            globalMouseY = hookStruct.pt.y;

            if (nCode >= 0 && wParam == (IntPtr)WM_LBUTTONDOWN)
            {
                if (bInLearning)
                {

                    mousePoints.Add(new MousePointInfo
                    {
                        MouseX = globalMouseX,
                        MouseY = globalMouseY,
                        Time = DateTime.Now,
                        TypeofAction = "DOWN"
                    });
  

                    iStepLearn++;
                }
            }
            if (nCode >= 0 && wParam == (IntPtr)WM_LBUTTONUP)
            {
                if (bInLearning)
                {

                    mousePoints.Add(new MousePointInfo
                    {
                        MouseX = globalMouseX,
                        MouseY = globalMouseY,
                        Time = DateTime.Now,
                        TypeofAction = "UP"
                    });

                    mousePoints[iStepLearn-1].Duration = (mousePoints[iStepLearn].Time - mousePoints[iStepLearn-1].Time).TotalSeconds;
                    mousePoints[iStepLearn].Duration = mousePoints[iStepLearn - 1].Duration;
                    iStepLearn++;
                }
            }
            return CallNextHookEx(hookId, nCode, wParam, lParam);
        }
     
        bool stopLearn = true;
        private void button2_Click(object sender, EventArgs e)
        {
            stopLearn = (stopLearn) ? false : true;
            if(stopLearn)
            {
                bInLearning = false;
                button2.BackColor = Color.LightGray;
                button1.Enabled = true;
                // Display a confirmation dialog
                DialogResult result = MessageBox.Show("Do you want to remove the last points?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                // Check the user's response
                if (result == DialogResult.Yes)
                {
                    if(mousePoints.Count >= 2)
                    {
                        mousePoints.RemoveRange(mousePoints.Count - 2, 2);
                    }
                }
             }
            else
            {
                button2.BackColor = Color.LightGreen;
                button1.Enabled = false;
                bInLearning = true;
                //this.WindowState = FormWindowState.Minimized;
            }
            //timer2.Start();
        }

        private void timer2_Tick(object sender, EventArgs e)
        {

            /*richTextBox1.Text = "";
            foreach(MousePointInfo NPI in mousePoints)
            {
                richTextBox1.Text = richTextBox1.Text + NPI.MouseX.ToString() + " " + NPI.MouseY.ToString() + " D: " + NPI.Duration.ToString() + " A: " + NPI.TypeofAction.ToString() + "\r\n";
            }*/
        }
    }
}