var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{

    string strURL = builder.Configuration.GetValue<string>("ServerURL");
    if (strURL != null && strURL != "")
    {
        c.AddServer(new Microsoft.OpenApi.Models.OpenApiServer()
        {
            Url = strURL
        });
    }

});

var app = builder.Build();

app.UseSwagger(p => p.SerializeAsV2 = true);
app.UseSwaggerUI(c =>
{
    c.SwaggerEndpoint("/swagger/v1/swagger.json", "PowerPoint Gen API V1");
    c.RoutePrefix = string.Empty;
});


app.UseRouting();

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseAuthorization();

app.MapControllers();

app.Run();
