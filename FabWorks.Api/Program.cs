using FabWorks.Api.Configuration;
using FabWorks.Api.Services;
using FabWorks.Core.Data;
using Microsoft.EntityFrameworkCore;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddDbContext<FabWorksDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("FabWorksDb")));
builder.Services.AddSingleton<FormProgramService>();

builder.Services.Configure<FileStorageOptions>(
    builder.Configuration.GetSection(FileStorageOptions.SectionName));
builder.Services.AddScoped<IFileStorageService, FileStorageService>();

var app = builder.Build();

app.UseDefaultFiles();
app.UseStaticFiles(new StaticFileOptions
{
    OnPrepareResponse = ctx =>
        ctx.Context.Response.Headers.Append("Cache-Control", "no-cache, no-store")
});
app.MapControllers();
app.Run();
