using System.Text.Json.Serialization;
using server.Data;
using server.Endpoints;

var builder = WebApplication.CreateBuilder(args);
var clientOrigin = builder.Configuration["Client:Origin"] ?? "http://localhost:3000";

builder.Services.AddOpenApi();

builder.Services.ConfigureHttpJsonOptions(options =>
{
	options.SerializerOptions.Converters.Add(new JsonStringEnumConverter());
});

builder.Services.AddSingleton<IRecordRepository, InMemoryRecordRepository>();

builder.Services.AddCors(options =>
{
	options.AddDefaultPolicy(policy =>
	{
		policy.WithOrigins(clientOrigin)
			.AllowAnyHeader()
			.AllowAnyMethod();
	});
});

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
	app.MapOpenApi();
}

app.UseCors();

app.MapRecordEndpoints();

app.Run();
