using Autofac;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    static class Program
    {
        private static IContainer _container; // autofac container
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Bootstrap();// wire up Autofac dependency graph for DI

            Application.Run( // inject form
                _container.Resolve<Form1>());
        }
        private static void Bootstrap()
        {
            var cb = new ContainerBuilder();

            ConnectionStringSettings _clientDbConn = ConfigurationManager.ConnectionStrings["SQLConn"];

            cb.RegisterType<PetRepository>()
                .As<IPetRepository>()
                .WithParameter(
                    "saConn",
                    _clientDbConn)
                .SingleInstance();

            cb.RegisterType<PetService>()
                .As<IPetService>()
                .SingleInstance();

            // register forms used
            cb.RegisterType<Form1>()
                .AsSelf()
                .InstancePerDependency();

            // build graph setting container 
            _container = cb.Build();
        }
    }
}
