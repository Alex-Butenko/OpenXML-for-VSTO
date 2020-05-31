using System;
using System.Reflection;
using System.Security;
using System.Security.Policy;

namespace OpenXmlForVsto.Word.BenchmarkAddin {
    public partial class ThisAddIn {
        void ThisAddIn_Startup(object sender, EventArgs e) {
            // This is to support big ranges. Otherwise OpenXML.dll may throw IsolatedStorageException
            Evidence newEvidence = new Evidence();
            newEvidence.AddHostEvidence(new Zone(SecurityZone.MyComputer));
            AppDomain.CurrentDomain
                .GetType()
                .GetField("_SecurityIdentity", BindingFlags.Instance | BindingFlags.NonPublic)?
                .SetValue(AppDomain.CurrentDomain, newEvidence);
        }

        void ThisAddIn_Shutdown(object sender, EventArgs e) {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        void InternalStartup() {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}