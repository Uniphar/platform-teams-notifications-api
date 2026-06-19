namespace Teams.Notifications.Api.Filters;

public class InvalidOperationExceptionFilter(ICustomEventTelemetryClient telemetry, ILogger<InvalidOperationExceptionFilter> logger) : IExceptionFilter
{
    public void OnException(ExceptionContext context)
    {
        if (context.Exception is not InvalidOperationException invalidOperationException) return;

        logger.LogError(invalidOperationException, "InvalidOperationByUser: {Message}", invalidOperationException.Message);

        telemetry.TrackEvent("InvalidOperationByUser",
            new()
            {
                ["Message"] = invalidOperationException.Message
            });

        context.Result = new ObjectResult(new
        {
            error = "Invalid Operation",
            message = invalidOperationException.Message
        })
        {
            StatusCode = StatusCodes.Status400BadRequest
        };

        context.ExceptionHandled = true;
    }
}