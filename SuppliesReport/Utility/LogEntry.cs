using SuppliesReport.EntityModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SuppliesReport.Utility
{
    public class LogEntry
    {
        public PRIZ Priz { get; set; }
        public LOG_P Log { get; set; }

        public LogEntry()
        {

        }

        public LogEntry(PRIZ p, LOG_P l)
        {
            Priz = p;
            Log = l;
        }

    }
}
