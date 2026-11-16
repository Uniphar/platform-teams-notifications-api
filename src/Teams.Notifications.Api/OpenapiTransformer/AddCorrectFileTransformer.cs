namespace Teams.Notifications.Api.OpenApiTransformer;

/// <summary>
///     Adds the correct file upload schema for multipart/form-data requests.
/// </summary>
public sealed class AddCorrectFileTransformer : IOpenApiOperationTransformer
{
    /// <summary>
    ///     Instead of the default file upload schema, we want to use the correct one for multipart/form-data
    ///     The standard one wants to use an IFormFile, which makes sense but Service Reference does not support that on the
    ///     client side.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="context"></param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    public Task TransformAsync(OpenApiOperation operation, OpenApiOperationTransformerContext context, CancellationToken cancellationToken)
    {
        if (operation.RequestBody?.Content == null || !operation.RequestBody.Content.TryGetValue("multipart/form-data", out var value)) return Task.CompletedTask;
        if (value.Schema is { Type: JsonSchemaType.Object, Properties: not null } && value.Schema.Properties.ContainsKey("file"))
        {
            operation.RequestBody.Content["multipart/form-data"] = new OpenApiMediaType
            {
                Encoding = new Dictionary<string, OpenApiEncoding>
                {
                    ["file"] = new()
                    {
                        Style = ParameterStyle.Form
                    }
                },
                Schema = new OpenApiSchema
                {
                    Properties = new Dictionary<string, IOpenApiSchema>
                    {
                        ["file"] = new OpenApiSchema
                        {
                            Type = JsonSchemaType.String,
                            Format = "binary"
                        }
                    }
                }
            };
        }

        return Task.CompletedTask;
    }
}