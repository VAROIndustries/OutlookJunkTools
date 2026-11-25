using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OutlookJunkTools.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkTools.AddIn
{
    /// <summary>The COM class Outlook loads. Wiring only; every entry point catches its own exceptions.</summary>
    [ComVisible(true)]
    [Guid("6F1C2B7A-3D54-4E0B-9A7E-2C1D8E5F4A10")]
    [ProgId("OutlookJunkTools.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private AddInState _state;
        private Commands _commands;
        private IncomingWatcher _watcher;

        public void OnConnection(object Application, int ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                var app = (Outlook.Application)Application;
                Log.Info("OnConnection mode=" + ConnectMode + " outlook=" + app.Version);
                _state = AddInState.Create(app);
                _commands = new Commands(_state);
                Log.Info("Block list: " + _state.List.Entries.Count + " entr" + (_state.List.Entries.Count == 1 ? "y" : "ies"));
            }
            catch (Exception ex) { Log.Error("OnConnection", ex); }
        }

        public void OnStartupComplete(ref Array custom)
        {
            try
            {
                if (_state == null) return;
                _watcher = new IncomingWatcher(_state);
                _watcher.Start();
                if (_state.LoadWarning != null)
                {
                    Log.Info(_state.LoadWarning);
                    MessageBox.Show(_state.LoadWarning, "Outlook Junk Tools", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex) { Log.Error("OnStartupComplete", ex); }
        }

        public void OnBeginShutdown(ref Array custom)
        {
            try { _watcher?.Stop(); } catch (Exception ex) { Log.Error("OnBeginShutdown", ex); }
        }

        public void OnDisconnection(int RemoveMode, ref Array custom)
        {
            try { _watcher?.Stop(); _watcher = null; _commands = null; _state = null; Log.Info("OnDisconnection mode=" + RemoveMode); }
            catch (Exception ex) { Log.Error("OnDisconnection", ex); }
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                string resource = RibbonID == "Microsoft.Outlook.Explorer" ? "Ribbon.xml"
                    : RibbonID == "Microsoft.Outlook.Mail.Read" ? "RibbonRead.xml" : null;
                if (resource == null) return null;
                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("OutlookJunkTools.AddIn." + resource))
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
            catch (Exception ex) { Log.Error("GetCustomUI " + RibbonID, ex); return null; }
        }

        // Ribbon callbacks. Outlook calls these by name through IDispatch.
        public void OnBlockSender(object control) => Run("Block Sender", () => _commands.Block(control, BlockKind.Address, false));
        public void OnBlockDomain(object control) => Run("Block Domain", () => _commands.Block(control, BlockKind.Domain, false));
        public void OnBlockAndJunk(object control) => Run("Block + Junk", () => _commands.Block(control, BlockKind.Address, true));
        public void OnJunk(object control) => Run("Junk", () => _commands.Junk(control));
        public void OnUnsubscribe(object control) => Run("Unsubscribe", () => _commands.Unsubscribe(control));
        public void OnManageList(object control) => Run("Block List", () => _commands.ManageList());

        private void Run(string name, Action action)
        {
            try
            {
                if (_commands == null) { Commands.Info("Outlook Junk Tools did not start correctly. See the log in %AppData%\\OutlookJunkTools."); return; }
                action();
            }
            catch (Exception ex)
            {
                Log.Error(name, ex);
                try { MessageBox.Show(name + " failed: " + ex.Message, "Outlook Junk Tools", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
            }
        }
    }
}
