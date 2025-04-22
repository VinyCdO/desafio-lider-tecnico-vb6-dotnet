var builder = WebApplication.CreateBuilder(args);

// Configuração de serviços
builder.Services.AddControllers();
builder.Services.AddScoped<IDividasRepository, DividasRepository>();
builder.Services.AddScoped<IDividasService, DividasService>();
builder.Services.AddScoped<INegociacoesRepository, NegociacoesRepository>();
builder.Services.AddScoped<INegociacoesService, NegociacoesService>();
builder.Services.AddAuthorization();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

builder.Logging.ClearProviders();
builder.Logging.AddConsole();

var app = builder.Build();

// Configuração do pipeline
app.UseHttpsRedirection();
app.UseAuthorization();
app.MapControllers();
app.UseSwagger();
app.UseSwaggerUI();

app.Run();