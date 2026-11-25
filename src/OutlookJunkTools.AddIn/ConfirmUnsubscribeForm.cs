using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookJunkTools.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkTools.AddIn
{
    /// <summary>Shows exactly what an unsubscribe will do and waits for a decision.</summary>
    internal sealed class ConfirmUnsubscribeForm : Form
    {
        private readonly AddInState _state;
        private readonly Outlook.MailItem _mail;
        private readonly IReadOnlyList<UnsubscribeOption> _options;
        private int _index;

        private readonly Label _method = new Label();
        private readonly TextBox _target = new TextBox();
        private readonly Label _status = new Label();
        private readonly Button _proceed = new Button();
        private readonly Button _browser = new Button();
        private readonly Button _cancel = new Button();

        public ConfirmUnsubscribeForm(AddInState state, Outlook.MailItem mail, string sender,
            IReadOnlyList<UnsubscribeOption> options, DkimVerdict dkim)
        {
            _state = state; _mail = mail; _options = options;

            Text = "Unsubscribe";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(560, 330);
            Font = SystemFonts.MessageBoxFont;

            var from = new Label { Text = "From: " + sender, Left = 14, Top = 14, Width = 532, AutoEllipsis = true };
            _method.SetBounds(14, 42, 532, 40);
            var targetLabel = new Label { Text = "Destination:", Left = 14, Top = 88, Width = 532 };
            _target.SetBounds(14, 108, 532, 62);
            _target.Multiline = true; _target.ReadOnly = true; _target.ScrollBars = ScrollBars.Vertical;

            string dkimText;
            switch (dkim)
            {
                case DkimVerdict.Pass:
                    dkimText = "Your mail server recorded DKIM as passed for this message."; break;
                case DkimVerdict.NotPass:
                    dkimText = "Your mail server did NOT record a DKIM pass. If this looks like spam, unsubscribing " +
                               "confirms your address is live. Block + Junk is the safer choice."; break;
                default:
                    dkimText = "No DKIM result was recorded for this message. If this looks like spam, unsubscribing " +
                               "confirms your address is live. Block + Junk is the safer choice."; break;
            }
            var dkimLabel = new Label { Text = dkimText, Left = 14, Top = 178, Width = 532, Height = 52 };
            if (dkim != DkimVerdict.Pass) dkimLabel.ForeColor = Color.Firebrick;

            _status.SetBounds(14, 234, 532, 40);

            _proceed.SetBounds(196, 286, 110, 30); _proceed.Text = "Proceed";
            _browser.SetBounds(312, 286, 150, 30); _browser.Text = "Open in browser";
            _cancel.SetBounds(468, 286, 78, 30); _cancel.Text = "Cancel";
            _proceed.Click += async (s, e) => await Proceed();
            _browser.Click += (s, e) => OpenBrowser();
            _cancel.Click += (s, e) => Close();
            CancelButton = _cancel;

            Controls.AddRange(new Control[] { from, _method, targetLabel, _target, dkimLabel, _status, _proceed, _browser, _cancel });
            ShowOption();
        }

        private UnsubscribeOption Current => _options[_index];

        private void ShowOption()
        {
            var o = Current;
            switch (o.Method)
            {
                case UnsubscribeMethod.OneClickPost:
                    _method.Text = "Method: one-click request. An unsubscribe request is sent straight to the sender's server. Nothing opens.";
                    _target.Text = o.Target.AbsoluteUri; break;
                case UnsubscribeMethod.Mailto:
                    _method.Text = "Method: email. An unsubscribe email is sent from the account that received this message.";
                    _target.Text = "To: " + o.MailAddress + Environment.NewLine + "Subject: " + o.MailSubject; break;
                case UnsubscribeMethod.HeaderLink:
                    _method.Text = "Method: web page from the message header. It opens in your browser for you to finish there.";
                    _target.Text = o.Target.AbsoluteUri; break;
                default:
                    _method.Text = "Method: unsubscribe link found in the message body. It opens in your browser for you to finish there.";
                    _target.Text = o.Target.AbsoluteUri; break;
            }
            bool isWeb = o.Method == UnsubscribeMethod.HeaderLink || o.Method == UnsubscribeMethod.BodyLink;
            _proceed.Text = isWeb ? "Open page" : "Proceed";
            _browser.Visible = o.Method == UnsubscribeMethod.OneClickPost;
        }

        private async Task Proceed()
        {
            var o = Current;
            SetBusy(true, "Working...");
            string error;
            try
            {
                switch (o.Method)
                {
                    case UnsubscribeMethod.OneClickPost:
                        error = await Task.Run(() => UnsubscribeRunner.PostOneClick(o.Target)); break;
                    case UnsubscribeMethod.Mailto:
                        error = UnsubscribeRunner.SendMailto(_state, _mail, o); break;
                    default:
                        error = UnsubscribeRunner.OpenInBrowser(o.Target); break;
                }
            }
            catch (Exception ex) { error = ex.Message; }

            Log.Info("Unsubscribe " + o.Method + " -> " + TargetForLog(o) + ": " + (error ?? "ok"));
            if (IsDisposed) return;

            if (error == null)
            {
                string done = o.Method == UnsubscribeMethod.OneClickPost ? "Unsubscribe request accepted by the sender's server."
                    : o.Method == UnsubscribeMethod.Mailto ? "Unsubscribe email sent."
                    : "Opened in your browser. Finish unsubscribing there.";
                MessageBox.Show(this, done, "Unsubscribe", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Close();
                return;
            }

            if (_index + 1 < _options.Count)
            {
                _index++;
                ShowOption();
                SetBusy(false, "That did not work: " + error + " Showing the next method this message offers.");
            }
            else
            {
                SetBusy(false, "That did not work: " + error);
            }
        }

        private void OpenBrowser()
        {
            var error = UnsubscribeRunner.OpenInBrowser(Current.Target);
            Log.Info("Unsubscribe browser -> " + TargetForLog(Current) + ": " + (error ?? "ok"));
            if (error == null) Close(); else _status.Text = "Could not open the browser: " + error;
        }

        private static string TargetForLog(UnsubscribeOption o) =>
            o.Method == UnsubscribeMethod.Mailto ? o.MailAddress : o.Target.Host;

        private void SetBusy(bool busy, string status)
        {
            _proceed.Enabled = !busy; _browser.Enabled = !busy;
            _status.Text = status;
            UseWaitCursor = busy;
        }
    }
}
