using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IBIMSGen.Penetration
{
    public class HostElement
    {
        public Element element;
        public RevitLinkInstance rli;

        public HostElement(Element element, RevitLinkInstance rli)
        {
            this.element = element;
            this.rli = rli;
        }
    }

}
