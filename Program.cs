using SageWebAPI;
var builder = WebApplication.CreateBuilder(args);

// Register controllers
builder.Services.AddControllers();

// Read the Sage settings from appsettings.json
var sqlInstance = builder.Configuration["Sage:SqlInstance"] ?? "";
var company = builder.Configuration["Sage:Company"] ?? "";
var username = builder.Configuration["Sage:Username"] ?? "";
var password = builder.Configuration["Sage:Password"] ?? "";

// Register SageConnector as a singleton — meaning one shared instance
// for the lifetime of the app. This is important because the Sage COM
// object is expensive to initialize and not thread-safe.
builder.Services.AddSingleton<SageConnector>(provider =>
{
    var connector = new SageConnector(sqlInstance, company, username, password);
    return connector;
});

var app = builder.Build();

app.UseHttpsRedirection();
app.MapControllers();

// Grab the SageConnector instance and initialize it on startup
// This proves the DLL is reachable before we accept any requests
var sage = app.Services.GetRequiredService<SageConnector>();
sage.Initialize();

app.Run();