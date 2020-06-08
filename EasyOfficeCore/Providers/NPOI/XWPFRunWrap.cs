using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyOffice.Providers.NPOI
{
    public class XWPFRunWrap
    {
        public XWPFRunWrap(XWPFRun run,int pos)
        {
            Run = run;
            Position = pos;
        }
        public XWPFRun Run { get; set; }
        public int Position { get; set; }
    }
}
