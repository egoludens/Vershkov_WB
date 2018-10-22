using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VershkovWB
{
    interface IProgressReportObserver
    {
        void NotifyAboutProgressReport(string progressReportUpdate);
    }
}
