using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.IO;
using System.Management.Instrumentation;
using System.Management;
using System.Security.AccessControl;
using System.Configuration;

namespace CustomActionInstall
{
    [RunInstaller(true)]
    public partial class CustoCreateFolder : System.Configuration.Install.Installer
    {
        public CustoCreateFolder()
        {
            InitializeComponent();
        }
        private int CreaCartella(string path)
        {
            try
            {
                DirectoryInfo dir = Directory.CreateDirectory(path);
                var nn = new FileSystemAccessRule(@"tsf\TFSSetup", FileSystemRights.FullControl, AccessControlType.Allow);
                var u = new System.Security.Principal.NTAccount("tsf", "tfssetup");
                var dirSec = new DirectorySecurity();
                dirSec.AddAccessRule(nn);
                Directory.SetAccessControl(path, dirSec);
                var user = System.IO.Directory.GetAccessControl(path).GetOwner(typeof(System.Security.Principal.NTAccount));
                return 0;
            }
            catch(Exception ex)
            {
                throw ex;
            }

        }
        private int CancellaCartella(string path)
        {
            try
            {
                string dirObject = "Win32_Directory.Name='" + path + "'";
                using (ManagementObject managementObject = new ManagementObject(dirObject))
                {
                    managementObject.Get();
                    ManagementBaseObject outParams = managementObject.InvokeMethod("Delete", null, null);

                    // ReturnValue should be 0, else failure
                    if (Convert.ToInt32(outParams.Properties["ReturnValue"].Value) != 0)
                    {
                        Console.Write("Errore");
                    }
                }
                return 0;

            }
            catch (Exception ex)
            {
                return 1; throw ex;
            }
        }

        [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Demand)]
        public override void Install(System.Collections.IDictionary stateSaver)
        {
            System.Diagnostics.Debugger.Launch();
            string rep = System.Configuration.ConfigurationManager.AppSettings["repDir"];
            string tmp = System.Configuration.ConfigurationManager.AppSettings["tmpDir"];
            CreaCartella(rep);
            CreaCartella(tmp);
            base.Install(stateSaver);
        }
        [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Demand)]
        public override void Uninstall(IDictionary savedState)
        {
            string rep = System.Configuration.ConfigurationManager.AppSettings["repDir"];
            string tmp = System.Configuration.ConfigurationManager.AppSettings["tmpDir"];
            if (Directory.Exists(rep))
            {
                CancellaCartella(rep);
            }
            if (Directory.Exists(tmp))
            {
                CancellaCartella(tmp);
            }
            base.Uninstall(savedState);
        }
        [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Demand)]
        public override void Rollback(IDictionary savedState)
        {
            string rep = System.Configuration.ConfigurationManager.AppSettings["repDir"];
            string tmp = System.Configuration.ConfigurationManager.AppSettings["tmpDir"];
            if (Directory.Exists(rep))
            {
                CancellaCartella(rep);
            }
            if (Directory.Exists(tmp))
            {
                CancellaCartella(tmp);
            }
            base.Rollback(savedState);
        }
    }
}
