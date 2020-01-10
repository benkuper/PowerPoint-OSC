using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace PowerPoint_OSC
{
    public partial class Ribbon1
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            localPortInput.Text = ThisAddIn.instance.localPort.ToString();
            remoteHostInput.Text = ThisAddIn.instance.remoteHost;
            remotePortInput.Text = ThisAddIn.instance.remotePort.ToString();
        }

        private void localPort_TextChanged(object sender, RibbonControlEventArgs e)
        {
            int local = ThisAddIn.instance.localPort;
            int remote = ThisAddIn.instance.remotePort;
            int.TryParse(localPortInput.Text, out local);
            int.TryParse(remotePortInput.Text, out remote);
            ThisAddIn.instance.setupOSC(local, remoteHostInput.Text, remote);
        }
    }
}
