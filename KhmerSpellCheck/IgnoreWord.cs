using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KhmerSpellCheck
{
    public class IgnoreWord : IEquatable<IgnoreWord>
    {
        public string selectedText { get; set; }
        public string khmerword { get; set; }
        public int startposition { get; set; }
        public string document { get; set; }
        public bool ignoreAll { get; set; }

        public bool Equals(IgnoreWord other)
        {
            if (this.selectedText == other.selectedText && this.khmerword == other.khmerword && this.startposition == other.startposition
                && this.document == other.document && this.ignoreAll == other.ignoreAll)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
