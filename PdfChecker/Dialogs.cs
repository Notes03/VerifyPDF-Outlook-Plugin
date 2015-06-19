using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;

namespace PdfChecker
{
    public static class Dialogs
    {

        /// <summary>
        /// Custom dialog wrapper
        /// </summary>
        /// <param name="text"></param>
        /// <param name="title"></param>
        /// <param name="button"></param>
        /// <param name="image"></param>
        /// <param name="icon"></param>
        /// <returns></returns>
        public static DialogResult ShowDialog(string text, string title, Buttons button, Image image, Icon icon)
        {
            Dialog dialog = new Dialog();
            dialog.Text = title;
            dialog.pictureBox.Image = image;
            dialog.Icon = icon;
            dialog.label.Text = text;

            switch (button)
            {
                case Buttons.OK:
                    dialog.btnYes.Visible = false;
                    dialog.btnNo.Visible = false;
                    dialog.btnOK.Visible = true;
                    break;
                case Buttons.YesNo:
                    dialog.btnYes.Visible = true;
                    dialog.btnNo.Visible = true;
                    dialog.btnOK.Visible = false;
                    break;
            }

            if (dialog.label.Height > 64)
            {
                dialog.Height = dialog.label.Top + dialog.label.Height + 78;
            }

            if (dialog.label.Width > 64)
            {
                dialog.Width = dialog.label.Left + dialog.label.Width + 20;
            }

            return dialog.ShowDialog();
            

        }

        /// <summary>
        /// Enum of possible button types
        /// </summary>
        public enum Buttons
        {
            OK,
            YesNo,
        }
    }
}
