using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Forms;

namespace QQAuto
{
    public partial class Form1 : Form
    {
        Thread thread = null;
        public Form1()
        {
            InitializeComponent();
        }
        [STAThread]
        private void button1_Click(object sender, EventArgs e)
        {
            if (thread == null)
            {
                thread = new Thread(QQAutoFile);
            }
            thread.Start();
            button1.Enabled = false;
            button2.Enabled = true;
        }

        private void QQAutoFile()
        {
            Process[] process = Process.GetProcessesByName("QQ");
            //获取根节点
            AutomationElement aeTop = AutomationElement.RootElement;
            foreach (Process p in process)
            {
                if (p.MainWindowHandle != null)
                {
                    try
                    {
                        AutomationElement aeForm = AutomationElement.FromHandle(p.MainWindowHandle);
                        if (aeForm != null)
                        {
                            ClickLeftMouse(aeForm);
                        }
                        else
                        {
                            Thread.Sleep(1000);
                        }
                    }
                    catch (Exception ex)
                    {
                        Thread.Sleep(1000);
                    }
                    
                }
            }
            QQAutoFile();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            thread.Suspend();
            button1.Enabled = true;
            button2.Enabled = false;
        }


        #region

        #region Import DLL
        /// <summary>
        /// Add mouse move event
        /// </summary>
        /// <param name="x">Move to specify x coordinate</param>
        /// <param name="y">Move to specify y coordinate</param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        extern static bool SetCursorPos(int x, int y);

        /// <summary>
        /// Mouse click event
        /// </summary>
        /// <param name="mouseEventFlag">MouseEventFlag </param>
        /// <param name="incrementX">X coordinate</param>
        /// <param name="incrementY">Y coordinate</param>
        /// <param name="data"></param>
        /// <param name="extraInfo"></param>
        [DllImport("user32.dll")]
        extern static void mouse_event(int mouseEventFlag, int incrementX, int incrementY, uint data, UIntPtr extraInfo);

        const int MOUSEEVENTF_MOVE = 0x0001;
        const int MOUSEEVENTF_LEFTDOWN = 0x0002;
        const int MOUSEEVENTF_LEFTUP = 0x0004;
        const int MOUSEEVENTF_RIGHTDOWN = 0x0008;
        const int MOUSEEVENTF_RIGHTUP = 0x0010;
        const int MOUSEEVENTF_MIDDLEDOWN = 0x0020;
        const int MOUSEEVENTF_MIDDLEUP = 0x0040;
        const int MOUSEEVENTF_ABSOLUTE = 0x8000;

        #endregion

        public static void ClickLeftMouse(AutomationElement element)
        {
            AutomationElement btn = element.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.NameProperty, "全接收"));

            AutomationElement btn2 = element.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.NameProperty, "全另存为"));

            if (btn != null)
            {
                Rect rect = btn.Current.BoundingRectangle;
                int IncrementX = (int)(element.Current.BoundingRectangle.BottomRight.X - (btn2.Current.BoundingRectangle.BottomRight.X - rect.BottomLeft.X) + rect.Width / 4);
                int IncrementY = (int)(element.Current.BoundingRectangle.BottomRight.Y - rect.Height / 2);

                //Make the cursor position to the element.
                SetCursorPos(IncrementX, IncrementY);

                //Make the left mouse down and up.
                mouse_event(MOUSEEVENTF_LEFTDOWN, IncrementX, IncrementY, 0, UIntPtr.Zero);
                mouse_event(MOUSEEVENTF_LEFTUP, IncrementX, IncrementY, 0, UIntPtr.Zero);
            }

        }

        #endregion

        /// <summary>
        /// Get the automation elemention of current form.
        /// </summary>
        /// <param name="processId">Process Id</param>
        /// <returns>Target element</returns>
        public static AutomationElement FindWindowByProcessId(int processId)
        {
            AutomationElement targetWindow = null;
            int count = 0;
            try
            {
                Process p = Process.GetProcessById(processId);
                targetWindow = AutomationElement.FromHandle(p.MainWindowHandle);
                return targetWindow;
            }
            catch (Exception ex)
            {
                count++;
                StringBuilder sb = new StringBuilder();
                string message = sb.AppendLine(string.Format("Target window is not existing.try #{0}", count)).ToString();
                if (count > 5)
                {
                    throw new InvalidProgramException(message, ex);
                }
                else
                {
                    return FindWindowByProcessId(processId);
                }
            }
        }

        /// <summary>
        /// Get the automation element by automation Id.
        /// </summary>
        /// <param name="windowName">Window name</param>
        /// <param name="automationId">Control automation Id</param>
        /// <returns>Automatin element searched by automation Id</returns>
        public static AutomationElement FindElementById(int processId, string automationId)
        {
            AutomationElement aeForm = FindWindowByProcessId(processId);
            AutomationElement tarFindElement = aeForm.FindFirst(TreeScope.Descendants,
            new PropertyCondition(AutomationElement.AutomationIdProperty, automationId));
            return tarFindElement;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Process[] process = Process.GetProcessesByName("QQAuto");
            foreach(Process p in process)
            {
                p.Kill();
            }
        }
    }
}
