using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Blueprint
{
    public class Project
    {
        public Vessel vessel;

        public string projectNumber { get; set; }
        public string serialNumber { get; set; }
        public string evapcoNumber { get; set; }
        public string drawingNumber { get; set; }
        public string revisionNumber { get; set; }
        public string approvalStatus { get; set; }
        public string customer { get; set; }
        public string poNumber { get; set; }

        public string vesselType { get; set; }
        public string designPressure { get; set; }
        public string maxTemperature { get; set; }
        public string minTemperature { get; set; }
        public string asmeEdition { get; set; }
        public string rtLong { get; set; }
        public string rtGirth { get; set; }
        public string testPressure { get; set; }
        public string caShell { get; set; }
        public string caHeads { get; set; }
        public string tMinShell { get; set; }
        public string tMinHeads { get; set; }
        public string tMinAfterForming { get; set; }
        public string paintSpecification { get; set; }
        public bool getsPWHT { get; set; }
        public bool getsFullVacuum { get; set; }
        public bool getsHydrotest { get; set; }
        public bool getsN2charge { get; set; }
        public bool getsTestClosures { get; set; }

        public string initials { get; set; }
        public string date { get; set; }


        public Project()
        {
            vessel = new Vessel();
        }


        public void generate_paperwork()
        {

        }
    }
}
