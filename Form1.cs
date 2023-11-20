using Microsoft.VisualBasic.Devices;
using System.Collections.Concurrent;
using System.ComponentModel;
using System.Diagnostics;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;
using static Clicker.Form1;
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

        // ConcurrentQueue to store information for each point
        private static ConcurrentQueue<MousePointInfo> mousePoints = new ConcurrentQueue<MousePointInfo>();

        private static int iStep = 0;
        //private static int iStepLearn = 0;

        private static bool bInLearning = false;
        private static bool bExecuteSequence = false;
        private DateTime targetDate;

        private static bool bHMI_Refresh_Cancel = false;
        private static bool bStepExec_Cancel = false;
        

        bool stop = true;
        private void button1_Click(object sender, EventArgs e)
        {
            stop = (stop) ? false : true;

            if (!stop)
            {
                bExecuteSequence = true;
                button1.BackColor = Color.LightGreen;
            }
            else
            {
                bExecuteSequence = false;
                button1.BackColor = Color.LightGray;
            }
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // Set the global mouse hook
            hookId = SetHook(llMouseProc);
            Task.Run(() => HMI_Refresh());
            Task.Run(() => Steps_Execute());
        }

        public void clickSimulator(Point p)
        {

            MousePointInfo currentPoint = mousePoints.Skip(iStep).FirstOrDefault();

            if (currentPoint != null)
            {
                SetCursorPos(currentPoint.MouseX, currentPoint.MouseY);

                if (currentPoint.TypeofAction == "DOWN")
                {
                    mouse_event((int)(MOUSEEVENTF_LEFTDOWN), p.X, p.Y, 0, 0);
                }
                if (currentPoint.TypeofAction == "UP")
                {
                    mouse_event((int)(MOUSEEVENTF_LEFTUP), p.X, p.Y, 0, 0);
                }
                Invoke(new Action(() => label1.Text = p.ToString()+ " Step : " + iStep.ToString()));

            }
        }

        private int tx;
        private int ty;
        private static double a { get; set; }
        private static double b { get; set; }
        private static int IntervalX;
        private static bool bGeneratePoints = false;
        private static List<int> IntervalGenerated = new List<int>();
        private static int iStepGenerated = 0;
        private static void ComputeLinearEquation(double x1, double y1, double x2, double y2)
        {
            // Compute the slope a
            a = (y2 - y1) / (x2 - x1);

            // Compute the offset b
            b = y1 - a * x1;
        }
        private static double Calculate(double x)
        {
            // Calculate y using the generated linear equation
            return a * x + b;
        }

        private void Steps_Execute()
        {
            int t = 0 ;
            while (!bStepExec_Cancel)
            {
                if (bExecuteSequence)
                {
                    MousePointInfo currentPoint = mousePoints.Skip(iStep).FirstOrDefault();
                    MousePointInfo nextPoint = mousePoints.Skip(iStep+1).FirstOrDefault();

                    //Main steps
                    if (currentPoint != null)
                    {
                        if (DateTime.Now >= targetDate)
                        {
                            targetDate = DateTime.Now.AddSeconds(currentPoint.Duration);
                            clickSimulator(new Point(MousePosition.X, MousePosition.Y));
                            if (nextPoint != null)
                            {
                                bGeneratePoints = true;
                                iStepGenerated = 0;
                            }
                            else
                            {
                                iStepGenerated = 0;
                                IntervalGenerated.Clear();
                            }
                        }
                        if(bGeneratePoints)
                        {
                            bGeneratePoints = false;
                            t = (int)currentPoint.Duration * 1000 / 50;
                            if(t>0)
                            {
                                if (currentPoint.MouseX < nextPoint.MouseX )
                                {
                                    IntervalX = (nextPoint.MouseX - currentPoint.MouseX ) / t;
                                }
                                else
                                {
                                    IntervalX = (nextPoint.MouseX - currentPoint.MouseX) / t;
                                    //IntervalX = (currentPoint.MouseX - nextPoint.MouseX ) / t;
                                }
                            
                                ComputeLinearEquation(currentPoint.MouseX, currentPoint.MouseY, nextPoint.MouseX, nextPoint.MouseY);
                                for(int i=1;i<= t ; i++)
                                {
                                    IntervalGenerated.Add((currentPoint.MouseX+(i*IntervalX)));
                                }
                            }
                        }
                        //Calculate(IntervalGenerated[iStepGenerated]);
                        if (IntervalGenerated.Count > 0)
                        {
                            if (iStepGenerated >= IntervalGenerated.Count - 1)
                            {
                                iStepGenerated = IntervalGenerated.Count - 1;
                                iStep++;
                            }
                            SetCursorPos(IntervalGenerated[iStepGenerated], (int)Calculate(IntervalGenerated[iStepGenerated]));
                            iStepGenerated++;
                        }
                        else
                        {
                            iStep++;
                        }
                    }
                    else
                    {
                        iStep = 0;
                        
                    }
                    

                }
                else
                {
                    iStep = 0;
                    iStepGenerated = 0;
                }

                Thread.Sleep(50);
            }
        }
        private void HMI_Refresh()
        {
            while (!bHMI_Refresh_Cancel)
            {
                Invoke(new Action(() => richTextBox1.Clear()));

                Invoke(new Action(() => progressBar1.Maximum = mousePoints.Count));
                if(iStep<= mousePoints.Count)
                {
                    Invoke(new Action(() => progressBar1.Value = iStep));
                }
                

                foreach (MousePointInfo NPI in mousePoints.ToArray())
                {
                    Invoke(new Action(() => richTextBox1.AppendText($"{NPI.MouseX} {NPI.MouseY} D: {NPI.Duration} A: {NPI.TypeofAction}\r\n")));
                }
              
                Thread.Sleep(200);
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            bHMI_Refresh_Cancel = true;
            bStepExec_Cancel = true;
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
                    mousePoints.Enqueue(new MousePointInfo
                    {
                        MouseX = globalMouseX,
                        MouseY = globalMouseY,
                        Time = DateTime.Now,
                        TypeofAction = "DOWN"
                    });

                }
            }
            if (nCode >= 0 && wParam == (IntPtr)WM_LBUTTONUP)
            {
                if (bInLearning)
                {
                    mousePoints.Enqueue(new MousePointInfo
                    {
                        MouseX = globalMouseX,
                        MouseY = globalMouseY,
                        Time = DateTime.Now,
                        TypeofAction = "UP"
                    });

                    // Convert ConcurrentQueue to a list
                    List<MousePointInfo> mousePointsList = mousePoints.ToList();

                    // Ensure there are at least two points
                    if (mousePointsList.Count >= 2)
                    {
                        // Access the last two points
                        MousePointInfo lastPoint = mousePointsList.Last();
                        MousePointInfo previousPoint = mousePointsList[mousePointsList.Count - 2];

                        //lastPoint.Duration = (lastPoint.Time - previousPoint.Time).TotalSeconds;
                        previousPoint.Duration = (lastPoint.Time - previousPoint.Time).TotalSeconds;
                        //previousPoint.Duration = lastPoint.Duration;
                    }

                }
               
            }
            return CallNextHookEx(hookId, nCode, wParam, lParam);
        }

        bool stopLearn = true;
        private void button2_Click(object sender, EventArgs e)
        {
            stopLearn = (stopLearn) ? false : true;
            if (stopLearn)
            {
                bInLearning = false;
                button2.BackColor = Color.LightGray;
                button1.Enabled = true;
                // Display a confirmation dialog
                DialogResult result = MessageBox.Show("Do you want to remove the last points?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                // Check the user's response
                if (result == DialogResult.Yes)
                {
                    if (mousePoints.Count >= 2)
                    {
                        mousePoints = new ConcurrentQueue<MousePointInfo>(mousePoints.Take(mousePoints.Count - 2));
                    }
                    //iStepLearn = mousePoints.Count;
                }
            }
            else
            {
                button2.BackColor = Color.LightGreen;
                button1.Enabled = false;
                bInLearning = true;
            }
        }

    }
}
