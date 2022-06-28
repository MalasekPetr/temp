using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using NBS.MailBox.BackEnd.Models;
using NBS.MailBox.BackEnd.Services;

namespace NBS.MailBox.BackEnd
{
    public class Startup
    {
        readonly string AllowedOrigins = "AllowedOrigins";

        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddControllers();
            services.Configure<MailSettings>(Configuration.GetSection("MailSettings"));
            services.Configure<ConfigStore>(Configuration.GetSection("ConfigStore")); 
            services.AddTransient<IMailService, Services.MailService>();

            services.AddCors(options =>
            {
                options.AddPolicy(name: AllowedOrigins,
                    builder =>
                    {
                        var origins = Configuration["AllowedOrigins"].Split(';');
                        builder.WithOrigins(origins)
                            .AllowAnyMethod()
                            .AllowCredentials()
                            .AllowAnyHeader();
                    });
            });

        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            app.UseRouting();
            app.UseCors(AllowedOrigins);
            app.UseAuthorization();
            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
            });
        }
    }
}
