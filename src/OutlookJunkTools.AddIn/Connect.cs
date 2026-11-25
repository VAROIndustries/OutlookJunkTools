using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkTools.AddIn
{
    [ComVisible(true)]
    [Guid("6F1C2B7A-3D54-4E0B-9A7E-2C1D8E5F4A10")]
    [ProgId("OutlookJunkTools.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private Outlook.Application _app;

        public void OnConnection(object Application, int ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _app = (Outlook.Application)Application;
                Log.Info("OnConnection mode=" + ConnectMode + " outlook=" + _app.Version);
            }
            catch (Exception ex) { Log.Error("OnConnection", ex); }
        }

        public void OnDisconnection(int RemoveMode, ref Array custom)
        {
            try { Log.Info("OnDisconnection mode=" + RemoveMode); _app = null; }
            catch (Exception ex) { Log.Error("OnDisconnection", ex); }
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom)
        {
            try { Log.Info("OnStartupComplete"); }
            catch (Exception ex) { Log.Error("OnStartupComplete", ex); }
        }

        public void OnBeginShutdown(ref Array custom) { }

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                Log.Info("GetCustomUI " + RibbonID);
                if (RibbonID != "Microsoft.Outlook.Explorer") return null;
                using (var stream = Assembly.GetExecutingAssembly()
                    .GetManifestResourceStream("OutlookJunkTools.AddIn.Ribbon.xml"))
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (Exception ex) { Log.Error("GetCustomUI", ex); return null; }
        }

        // Ribbon callback.
        public void OnPing(object control)
        {
            try { Log.Info("Ping"); }
            catch (Exception ex) { Log.Error("OnPing", ex); }
        }
    }
}
