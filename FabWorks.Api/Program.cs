using FabWorks.Api.Services;
using FabWorks.Core.Data;
using Microsoft.EntityFrameworkCore;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddDbContext<FabWorksDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("FabWorksDb")));
builder.Services.AddSingleton<FormProgramService>();

var app = builder.Build();

app.MapControllers();
app.Run();
