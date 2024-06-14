using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.Windows;
using WPFStarter.Models.Data;

namespace WPFStarter
{
    public partial class App : Application
    {
        private readonly IHost host;

        public App()
        {
            host = Host.CreateDefaultBuilder()
                .ConfigureServices((context, services) =>
                {
                    var connectionString = context.Configuration.GetConnectionString("DefaultConnection");
                    services.AddDbContext<ApplicationDbContext>(options =>
                        options.UseSqlServer(connectionString));
                    services.AddSingleton<MainWindow>();
                })
                .Build();
        }

        protected override async void OnStartup(StartupEventArgs e)
        {
            await host.StartAsync();
            var mainWindow = host.Services.GetRequiredService<MainWindow>();
            mainWindow.Show();
            base.OnStartup(e);
        }

        protected override async void OnExit(ExitEventArgs e)
        {
            await host.StopAsync();
            host.Dispose();
            base.OnExit(e);
        }
    }
}