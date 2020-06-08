using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyOffice.Models.Word
{
    public class RunCollection : IEnumerable<Run>
    {
        private List<Run> runs = new List<Run>();
        public IEnumerator<Run> GetEnumerator()
        {
            return runs.GetEnumerator();
        }
        public RunCollection CreateRun(Action<Run> action)
        {
            Run run = new Run();
            runs.Add(run);
            action(run);
            return this;
        }
        public void AppendRun(Run run)
        {
            runs.Add(run);
        }
        IEnumerator IEnumerable.GetEnumerator()
        {
            return runs.GetEnumerator();
        }
    }
}
