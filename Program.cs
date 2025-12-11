using Microsoft.AspNetCore.Mvc.RazorPages;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddAuthentication("CookieAuth")
    .AddCookie("CookieAuth", options =>
    {
        options.Cookie.Name = "SurveyDashboard.Auth";
        options.LoginPath = "/Login";
        options.LogoutPath = "/Logout"; // We'll handle logout via handler in Login or separate
        options.ExpireTimeSpan = TimeSpan.FromHours(8); // Keep session for a work day
    });

builder.Services.AddAuthorization();

builder.Services.AddRazorPages(options =>
{
    // Future filters or conventions can be configured here.
});

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

app.MapRazorPages();

app.Run();