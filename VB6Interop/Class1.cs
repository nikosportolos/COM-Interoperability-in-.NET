using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace VB6Interop
{
    [Guid("6DF9A48B-E725-4735-955A-2BAC5439A2BA")]
    [ComVisible(true)]
    public interface IVB6Interop
    {
        [DispId(1)]
        void SampleMethod(string Message);
    }

    [Guid("7E6D6368-0033-49F6-9FE3-B2D409572869")]
    [ProgId("VB6Interop")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IVB6Interop))]
    [ComVisible(true)]
    public class VB6Interop : IVB6Interop
    {
        public void SampleMethod(string Message)
        {
            try
            {
                //MessageBox.Show(Message);
                FireSampleEvent("I received the message: " + Message);
            }
            catch (Exception ex)
            {                
                throw new Exception("Exception occured in SampleMethod(): ", ex);
            }
        }

        // Create delegate
        public delegate void SampleEventHandler(string Message);

        // Create event
        public new event SampleEventHandler SampleEvent = null;

        // Create event trigger
        public void FireSampleEvent(string Message)
        {
            try
            {
                if (SampleEvent != null)
                    SampleEvent(Message);
            }
            catch (Exception ex)
            {
                throw new Exception("Exception occured in FireSampleEvent(): ", ex);
            }
        }
    }

}
