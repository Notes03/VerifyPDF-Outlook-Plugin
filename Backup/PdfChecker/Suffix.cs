using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace PdfChecker
{
    enum Suffix
    {
        [Description("pdf")]
        PDF,
        [Description("zip")]
        ZIP,
        [Description("gz")]
        GZ,
        [Description("tar")]
        TAR,
        [Description("7z")]
        SEVENZIP,
        [Description("rar")]
        RAR
    }
}
