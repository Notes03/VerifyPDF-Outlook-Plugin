using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PdfChecker
{
    sealed public class ConverImage : System.Windows.Forms.AxHost
    {
        private ConverImage() : base(null)
        {
        }

        public static stdole.IPictureDisp Convert(System.Drawing.Image image)
        {
            System.Windows.Forms.ImageList list = new System.Windows.Forms.ImageList();
            list.Images.Add(image);
            stdole.IPictureDisp picture = (stdole.IPictureDisp) System.Windows.Forms.AxHost.GetIPictureDispFromPicture(list.Images[0]);
            return picture;
        }
    }
}
