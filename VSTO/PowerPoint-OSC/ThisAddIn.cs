using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using SharpOSC;
using System.Runtime.InteropServices;

namespace PowerPoint_OSC
{
    public partial class ThisAddIn
    {
        UDPListener listener;
        UDPSender sender;

        public int localPort;
        public String remoteHost;
        public int remotePort;

        public static ThisAddIn instance;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            instance = this;
            setupOSC(35550,"127.0.0.1",35551);

            Application.SlideShowNextSlide += onNextSlide;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if(listener != null) listener.Close();
        }

        public void setupOSC(int _localPort, string _remoteHost, int _remotePort)
        {
            localPort = _localPort;
            remoteHost = _remoteHost;
            remotePort = _remotePort;
            if (listener != null) listener.Close();

            HandleOscPacket callback = delegate (OscPacket packet)
            {
                var messageReceived = (OscMessage)packet;
                if(messageReceived != null) Console.WriteLine("Received a message! "+messageReceived.Address);
            };
              
            listener = new UDPListener(localPort, processMessage);

            sender = new UDPSender(remoteHost,remotePort);

            sendCurrentSlide();
            sendTotalSlides();
        }

       
        void processMessage(OscPacket packet)
        {
            var msg = (OscMessage)packet;
            if (msg.Address == "/next")
            {
                try
                {
                    Application.ActivePresentation.SlideShowWindow.View.Next();
                }
                catch (COMException)
                {
                    // Completely unknown error
                }

            }
            else if (msg.Address == "/previous")
            {
                try
                {
                    Application.ActivePresentation.SlideShowWindow.View.Previous();
                }
                catch (COMException)
                {
                    // Completely unknown error
                }
            }
            else if (msg.Address == "/slide")
            {
                if(msg.Arguments.Count >= 1)
                {
                    int page = (int)msg.Arguments[0];
                    try
                    {
                        Application.ActivePresentation.SlideShowWindow.View.GotoSlide(page);
                    }
                    catch (COMException)
                    {
                        // Completely unknown error
                    }
                }
               
            }
        }

        //Events
        private void onNextSlide(SlideShowWindow Wn)
        {
            // slide changed, send data
            sendCurrentSlide();
            sendTotalSlides();
        }

        void sendCurrentSlide()
        {
            if(Application.Presentations.Count == 0) return;
            OscMessage m = new OscMessage("/currentSlide", Application.ActivePresentation.SlideShowWindow.View.CurrentShowPosition);
            sender.Send(m);
        }

        void sendTotalSlides()
        {
            if (Application.Presentations.Count == 0) return;
            OscMessage m = new OscMessage("/totalSlides", Application.ActivePresentation.Slides.Count);
            sender.Send(m);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
