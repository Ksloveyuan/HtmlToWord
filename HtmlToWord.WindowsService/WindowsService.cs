using System.ServiceModel;
using System.ServiceProcess;
using HtmlToWord.Service;

namespace HtmlToWord.WindowsService
{
    public partial class WindowsService : ServiceBase
    {
        private ServiceHost _serviceHost;

        public WindowsService()
        {
            this.InitializeComponent();
            this.ServiceName = "example Export Service";
        }

        protected override void OnStart(string[] args)
        {
            this._serviceHost?.Close();

            this._serviceHost = new ServiceHost(typeof(ConvertService));

            this._serviceHost.Open();
        }

        protected override void OnStop()
        {
            if (this._serviceHost != null)
            {
                this._serviceHost.Close();
                this._serviceHost = null;
            }
        }
    }
}