using System;
using System.Runtime.CompilerServices;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Threading;
using HtmlToWord.Core;
using HtmlToWord.Service;
using log4net;

namespace HtmlToWord.ConsoleHost
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            log4net.Config.XmlConfigurator.Configure();

            while (true)
            {
                var logger = new Logger();
                var host = new WebServiceHost(typeof(ConvertService));
                try
                {
                    host.Open();
                    logger.Info("Service started.");
                    Console.ReadLine();

                    host.Close();
                }
                catch (CommunicationException cex)
                {
                    logger.Error("An exception occurred: {0}", cex.Message);
                    host.Abort();
                }
                finally
                {
                    Thread.Sleep(TimeSpan.FromSeconds(1));
                }
            }
        }
    }
}