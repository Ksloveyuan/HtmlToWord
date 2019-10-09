using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace HtmlToWord.WindowsService
{
    // Provide the ProjectInstaller class which allows 
    // the service to be installed by the Installutil.exe tool
    [RunInstaller(true)]
    public class ProjectInstaller : Installer
    {
        private readonly ServiceProcessInstaller _process;
        private readonly ServiceInstaller _service;

        public ProjectInstaller()
        {
            this._process = new ServiceProcessInstaller {Account = ServiceAccount.LocalSystem };
            this._service = new ServiceInstaller {ServiceName = "example_Export_Word_Service", DisplayName = "example Export Word Service", StartType = ServiceStartMode.Automatic};

            this.Installers.Add(this._process);
            this.Installers.Add(this._service);
        }
        public override void Commit(IDictionary savedState)
        {
            base.Commit(savedState);
            var sc = new ServiceController("example_Export_Word_Service");
            if (sc.Status.Equals(ServiceControllerStatus.Stopped))
            {
                sc.Start();
            }
        }
    }
}
