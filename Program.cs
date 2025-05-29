using Rotativa.AspNetCore;
using System.IO;       
using DataProfile.DataAccessLayer;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddControllersWithViews();

builder.Services.AddScoped<IProfileRepository, ProfileRepository>(); 

var app = builder.Build(); 

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}
app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();
app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
     pattern: "{controller=Home}/{action=Index}/{id?}");

// Konfigurasi Rotativa menggunakan 'app.Environment'
var env = app.Environment; 
Rotativa.AspNetCore.RotativaConfiguration.Setup(env.WebRootPath, "Rotativa");

app.Run();