using Syncfusion.Licensing;
using System.ServiceProcess;

namespace CmaFileWatcherService
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            SyncfusionLicenseProvider.RegisterLicense("Ngo9BigBOggjHTQxAR8/V1NDaF5cWWtCf1FpRmJGdld5fUVHYVZUTXxaS00DNHVRdkdnWH1ecnVVRmJeVUN/WUA=");
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new CmaFileWatcherService()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
