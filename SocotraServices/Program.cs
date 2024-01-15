using log4net;
using log4net.Config;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddHttpClient();

var log4NetConfigFile = "log4net.config";
var logRepository = LogManager.GetRepository(System.Reflection.Assembly.GetEntryAssembly());
XmlConfigurator.Configure(logRepository, new FileInfo(log4NetConfigFile));
builder.Services.AddSingleton<ILog>(LogManager.GetLogger(typeof(Program)));


builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowSpecificOrigin", builder =>
    {
        builder.AllowAnyOrigin() // Allow requests from any origin
                    .AllowAnyHeader()
                    .AllowAnyMethod();
    });
});

//builder.Services.AddResponseCompression(options =>
//{
//    options.EnableForHttps = true; // Enable compression for HTTPS requests
//    options.MimeTypes = new[] { "application/json", "text/plain" }; // MIME types to compress
//});


var app = builder.Build();

//app.UseResponseCompression();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
