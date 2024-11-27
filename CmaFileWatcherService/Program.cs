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
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new CmaFileWatcherService()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
