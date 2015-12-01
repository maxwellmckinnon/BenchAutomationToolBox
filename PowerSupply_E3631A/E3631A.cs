using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ivi.Visa.Interop;

namespace PowerSupply_E3631A
{
    public class E3631A
    {
        ResourceManager rMgr = new ResourceManagerClass();
        FormattedIO488 src = new FormattedIO488Class();
        string srcAddress = "GPIB::03";

        public E3631A()
        {
            connectGPIB();
        }
        
        public void setVoltage(double voltage, string supply, double current = 1)
        {
            //string supply = "P6V", "P25V", or "N25V"
            src.WriteString("APPLy " + supply + ", " + voltage.ToString("F4") + ", " + current.ToString());
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
