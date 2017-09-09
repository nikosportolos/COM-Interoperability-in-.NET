using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace InterprocessCommunication
{
    class WindowMessaging : NativeWindow, IDisposable
    {

        #region WIN32 Declarations

		// API function to create custom system-wide Messages
		[DllImport("user32.dll")]
		public static extern IntPtr RegisterWindowMessage (String lpString);

		// API function to find window based on WindowName and class
		[DllImport("user32.dll")]
		public static extern IntPtr FindWindow (string lpClassName, string lpWindowName);

		// API function to send async. message to target application
		[DllImport("user32.dll")]
		public static extern IntPtr PostMessage (IntPtr hwnd, IntPtr wMsg, Int32 wParam, Int32 lParam);

		#endregion
        
		#region Declarations
		private bool disposed = false;
        
        // Message
        private string MSG_HELLO_REQ = "Hello VB6";
        private string MSG_HELLO_RESP = "Hello C#";

        // Window Title
        private string VB6WindowTitle = "VB6Messaging";

        // Message ID
        private IntPtr InMessageID;
        private IntPtr OutMessageID;

		#endregion

        #region Public Properties

        public IntPtr MyWindowHandle
        {
            get { return FindWindow(null, VB6WindowTitle); }
        }

        public IntPtr HELLO_REQ_MessageId
        {
            get
            {
                if (IntPtr.Zero.Equals(OutMessageID))
                {
                    OutMessageID = RegisterWindowMessage(MSG_HELLO_REQ);
                }

                return OutMessageID;
            }
        }

        public IntPtr SHUTDOWN_RESP_MessageId
        {
            get
            {
                if (IntPtr.Zero.Equals(InMessageID))
                    InMessageID = RegisterWindowMessage(MSG_HELLO_RESP);

                return InMessageID;
            }
        }

        #endregion

        #region Window Messaging

        // Function which implements custom message handling. Either way, then pass Message Handling through to base handler (i.e MainForm)
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == InMessageID.ToInt32())
            {
                MessageBox.Show("Incoming Windows Message", "WindowsMessaging");
            }

            base.WndProc(ref m);
        }

        //Function to find WindowHandle of VB6 app, and send async message through.
        public void SendMessage()
        {
            IntPtr hwndTarget = this.MyWindowHandle;
            IntPtr MessageId = OutMessageID;

            if (!IntPtr.Zero.Equals(hwndTarget))
            {
                PostMessage(hwndTarget, MessageId, 0, 0);
            }
        }

        #endregion

        #region Constructor and Dispose Methods

		public WindowMessaging()
		{
			//Creating Window based on globally known name, and create handle so we can start listening to Window Messages.
			CreateParams Params  = new CreateParams();
			Params.Caption = "CS";
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
			if(!this.disposed)
			{
				// If disposing equals true, dispose all managed and unmanaged resources.
				if(disposing)
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
