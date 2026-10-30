using Microsoft.Teams.Apps;
using NotesTeammate;

var builder = WebApplication.CreateBuilder(args);

// Teams SDK: validates incoming requests against the blueprint (the "AzureAd" section) and gets tokens for each
// hired instance's agent user.
builder.Services.AddTeamsBotApplication();
builder.Services.AddHireTenants(builder.Configuration); // hires can be in any customer tenant
builder.Services.AddHttpClient();
builder.Services.AddSingleton(TimeProvider.System);
builder.Services.AddSingleton<IInstanceStore>(new FileInstanceStore(
    Path.Combine(builder.Environment.ContentRootPath, builder.Configuration["Storage:Folder"] ?? "data")));
builder.Services.AddSingleton<Teammate>();
builder.Services.AddSingleton<WorkIq>();
builder.Services.AddSingleton<AgentDirectory>();
builder.Services.AddSingleton<AnswerFinder>();
builder.Services.AddSingleton<TeamsRoutes>();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseActivityLog(); // logs each activity's type and name, to watch lifecycle events arrive
}

// POST /api/messages: the endpoint set on the blueprint. Every hired instance is served from here.
var teams = app.UseTeamsBotApplication();
app.Services.GetRequiredService<TeamsRoutes>().Map(teams);

app.Run();
