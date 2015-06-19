using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Core;
using System.Reflection;
using System.ComponentModel;
using System.Windows.Forms;

namespace PdfChecker
{
    public partial class ThisAddIn
    {
        private CommandBarButton toolbarButton;

        private void onToolbarButtonClick(CommandBarButton button, ref bool cancel)
        {
            if (this.Application.ActiveExplorer().Selection.Count > 0)
            {
                Object selectedObject = this.Application.ActiveExplorer().Selection[1];
                if (selectedObject is MailItem)
                {
                    MailItem mail = (selectedObject as MailItem);
                    if (mail.Attachments != null)
                    {
                        Attachments attachments = mail.Attachments;
                        if (TestAttachments(attachments))
                        {
                            SendMail(mail);
                        }
                    }
                }
            }
        }

        private bool TestAttachments(Attachments attachments)
        {
            string text;
            foreach (Attachment attachment in attachments)
            {
                foreach (Suffix suffix in Enum.GetValues(typeof(Suffix)))
                {
                    if (attachment.FileName.ToLower().EndsWith(GetDescription(suffix)))
                    {
                        System.Diagnostics.Debug.Print("text check");
                        text = "Möchten Sie den Anhang dieser E-Mail auf gefähliche PDF-Dateien überprüfen lassen?"
                            + "\n\nWenn Sie fortfahren, wird der Anhang diese E-Mail an check@verifypdf.com weitergeleitet";
                        DialogResult result = MessageBox.Show(text, "PDF Check", MessageBoxButtons.YesNo);
                        if(result == DialogResult.Yes)
                        {
                            return true;
                        }
                        else return false;
                    }
                }
                text = "Der Anhang dieser E-Mail besitzt kein Unterstütztes Dateiformat."
                        + "\n\nSollten sich dennoch PDF-Dateien darin befinden, so können Sie diese unter www.verifyPDF.com überprüfen."; 
                MessageBox.Show(text, "PDF Check");
                return false;
            }
            text = "Diese E-Mail besitzt keinen Anhang."
                + "\n\nWeitere Informationen zum Angebot von VerifyPDF finden Sie unter www.verifyPDF.com.";
            MessageBox.Show(text, "PDF Check");
            return false;
        }

        private void SendMail(MailItem mail)
        {
            MailItem newMail = mail.Copy();
            newMail.Subject = "VerifyPDF (" + mail.Subject + ")";
            newMail.Body = "";
            //newMail.To = "check@verifypdf.com";
            newMail.To = "thomas.hanka@hs-augsburg.de";
            ((_MailItem)newMail).Send();
            MessageBox.Show("Der Anhang wurde weitergeleitet.\n\nSie erhalten in Kürze die Auswertung.", "PDF Check");
        }

        private string GetDescription(Enum value)
        {
            FieldInfo info = value.GetType().GetField(value.ToString());
            DescriptionAttribute[] attributes = (DescriptionAttribute[]) info.GetCustomAttributes(typeof(DescriptionAttribute), false);
            if (attributes != null && attributes.Length > 0)
            {
                return attributes[0].Description;
            }
            else return value.ToString();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            CommandBars commandBars = this.Application.ActiveExplorer().CommandBars;

            try
            {
                this.toolbarButton = (CommandBarButton)commandBars["Standard"].Controls["PDF Check"];

            }
            catch (System.Exception)
            {
                this.toolbarButton =
                    (CommandBarButton)commandBars["Standard"].Controls.Add(1,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                this.toolbarButton.Caption = "PDF Check";
                this.toolbarButton.Picture = PdfChecker.ConverImage.Convert(Properties.Resources.iconOl);
                this.toolbarButton.Mask = PdfChecker.ConverImage.Convert(Properties.Resources.iconOlMask);
                this.toolbarButton.Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonIcon;
            }
            this.toolbarButton.Tag = "PDFCheck";
            this.toolbarButton.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(this.onToolbarButtonClick);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.toolbarButton.Delete(System.Reflection.Missing.Value);
            this.toolbarButton = null;
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
