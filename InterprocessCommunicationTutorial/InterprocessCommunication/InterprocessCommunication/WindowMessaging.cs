using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Windows.Forms;

namespace InterprocessCommunication
{
    public class WindowMessaging : NativeWindow, IDisposable
    {
        #region WIN32 Declarations

        // API function to create custom system-wide Messages
        [DllImport("user32.dll")]
        public static extern IntPtr RegisterWindowMessage(String lpString);

        // API function to find window based on WindowName and class
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        // API function to send async. message to target application
        [DllImport("user32.dll")]
        public static extern IntPtr PostMessage(IntPtr hwnd, IntPtr wMsg, Int32 wParam, Int32 lParam);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        /// <summary>
        /// The class exposes Windows APIs to be used in this code sample.
        /// </summary>
        [SuppressUnmanagedCodeSecurity]
        internal class NativeMethod
        {
            /// <summary>
            /// The FindWindow function retrieves a handle to the top-level window whose class name and window name match the specified strings. 
            /// This function does not search child windows. This function does not perform a case-sensitive search.
            /// </summary>
            /// <param name="lpClassName">Class name</param>
            /// <param name="lpWindowName">Window caption</param>
            /// <returns></returns>
            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        }
        #endregion

        #region Declarations
        private bool disposed = false;

        // Message
        private string MSG_HELLO_REQUEST = "MSG_HELLO_REQUEST";
        private string MSG_HELLO_RESPONSE = "MSG_HELLO_RESPONSE";
        private string MSG_VB6_TO_CSHARP = "MCL_VB6_TO_C#";
        private string MSG_CSHARP_TO_VB6 = "MCL_C#_TO_VB6";

        //// Window Title
        private string VB6WindowTitle = "VB6-MessageServer";
        private string CSWindowTitle = "C#-MessageServer";

        //// Message IDs
        private IntPtr _HELLO_REQ_MessageId;
        private IntPtr _HELLO_RESP_MessageId;
        private IntPtr VB6_TO_CSHARPMessageId;
        private IntPtr CSHARP_TO_VB6MessageId;

        #endregion

        #region Public Properties

        public IntPtr VB6_WindowHandle
        {
            get { return FindWindow(null, VB6WindowTitle); }
        }

        public IntPtr HELLO_REQ_MessageId
        {
            get
            {
                if (IntPtr.Zero.Equals(_HELLO_REQ_MessageId))
                {
                    _HELLO_REQ_MessageId = RegisterWindowMessage(MSG_HELLO_REQUEST);
                }

                return _HELLO_REQ_MessageId;
            }
        }

        public IntPtr HELLO_RESP_MessageId
        {
            get
            {
                if (IntPtr.Zero.Equals(_HELLO_RESP_MessageId))
                {
                    _HELLO_RESP_MessageId = RegisterWindowMessage(MSG_HELLO_RESPONSE);
                }

                return _HELLO_RESP_MessageId;
            }
        }

        public IntPtr VB6_TO_CSHARP_MessageId
        {
            get
            {
                if (IntPtr.Zero.Equals(VB6_TO_CSHARPMessageId))
                {
                    VB6_TO_CSHARPMessageId = RegisterWindowMessage(MSG_VB6_TO_CSHARP);
                }

                return VB6_TO_CSHARPMessageId;
            }
        }

        public IntPtr CSHARP_TO_VB6_MessageId
        {
            get
            {
                if (IntPtr.Zero.Equals(CSHARP_TO_VB6MessageId))
                {
                    CSHARP_TO_VB6MessageId = RegisterWindowMessage(MSG_CSHARP_TO_VB6);
                }

                return CSHARP_TO_VB6MessageId;
            }
        }

        #endregion

        #region Window Messaging

        // Function which implements custom message handling. Either way, then pass Message Handling through to base handler (i.e MainForm)
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == HELLO_RESP_MessageId.ToInt32())
            {
                MessageBox.Show("Incoming Windows Message", "Interprocess Communication");
            }
            else 
            if (m.Msg == VB6_TO_CSHARP_MessageId.ToInt32())
            {
                MessageBox.Show("Windows Message from VB6 environment", "C#Messaging");
            }

            base.WndProc(ref m);
        }

        public string GetActiveWindowTitle()
        {
            const int nChars = 256;
            StringBuilder Buff = new StringBuilder(nChars);
            IntPtr handle = GetForegroundWindow();

            if (GetWindowText(handle, Buff, nChars) > 0)
            {
                return Buff.ToString();
            }

            return null;
        }

        // Function to find WindowHandle of VB6 app, and send async message through.
        public void SendMessage()
        {
            IntPtr hwndTarget = this.VB6_WindowHandle;
            IntPtr MessageId = HELLO_REQ_MessageId;

            // Find the target window handle.
            IntPtr hTargetWnd = NativeMethod.FindWindow(null, VB6WindowTitle);
            if (hTargetWnd == IntPtr.Zero)
            {
                MessageBox.Show("VB6 Window not found!");
                return;
            }

            // Send message to VB6
            PostMessage(hwndTarget, MessageId, 0, 0);
        }

        #endregion

        #region Constructor and Dispose Methods

        public WindowMessaging()
        {
            //Creating Window based on globally known name, and create handle so we can start listening to Window Messages.
            CreateParams Params = new CreateParams();
            Params.Caption = CSWindowTitle;
            this.CreateHandle(Params);
        }

        public void Dispose()
        {
            Dispose(true);
            // This object will be cleaned up by the Dispose method. Therefore, you should call GC. SupressFinalize to take this 
            // object off the finalization queue and prevent finalization code for this object from executing a second time.
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (!this.disposed)
            {
                // If disposing equals true, dispose all managed and unmanaged resources.
                if (disposing)
                {

                }

                if (!this.Handle.Equals(IntPtr.Zero))
                {
                    this.ReleaseHandle();
                }
            }
            disposed = true;
        }

        ~WindowMessaging()
        {
            // Do not re-create Dispose clean-up code here. Calling Dispose(false) is  
            // optimal in terms of readability and maintainability.
            Dispose(false);
        }
        #endregion
    }
}
