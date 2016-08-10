using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ivi.Visa.Interop;

namespace PowerSupply
{
    public class KikusuiPBZ20_20
    {
        ResourceManager rMgr = new ResourceManagerClass();
        FormattedIO488 src = new FormattedIO488Class();
        string srcAddress = "GPIB::07";

        public KikusuiPBZ20_20(string addr)
        {
            setGPIBAddress(addr);
            connectGPIB();
        }
        public KikusuiPBZ20_20()
        {
            connectGPIB();
        }

        public void setDCVoltage(double voltage)
        {
            src.WriteString("VOLT " + voltage.ToString("F3"));
        }

        public void setACVoltage(double voltage)
        {
            src.WriteString("VOLT:AC " + voltage.ToString("F3"));
        }

        public void setACFrequency(double freq)
        {
            src.WriteString("FREQ " + freq.ToString("F6"));
        }

        public void setGPIBAddress(string addr)
        {
            this.srcAddress = addr;
        }

        public void connectGPIB()
        {
            src.IO = (IMessage)rMgr.Open(srcAddress);
        }
    }
}
