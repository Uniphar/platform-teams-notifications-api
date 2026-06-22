namespace Teams.Notifications.Api.OpenApiTransformer;

/// <summary>
///     Transforms an OpenAPI document to annotate GET operations with Adaptive Card schema documentation.
/// </summary>
/// <remarks>
///     This transformer updates the OpenAPI document so that GET operations returning JSON responses are
///     marked with external documentation referencing the Adaptive Card schema. This enables consumers of the API
///     specification to discover and understand the structure of Adaptive Card responses. The transformer is typically
///     used
///     in scenarios where API responses conform to the Adaptive Card format and additional schema context is beneficial
///     for
///     client generation or documentation tools.
/// </remarks>
public sealed class AddAdaptiveCardDocsTransformer : IOpenApiDocumentTransformer
{
    /// <summary>
    ///     Transforms the specified OpenAPI document by updating the schema and external documentation for all GET
    ///     operations with a 200 response.
    /// </summary>
    /// <remarks>
    ///     For each GET operation with a 200 response, the method sets the response schema type to
    ///     "object" and attaches external documentation referencing Adaptive Cards. The input document is modified in
    ///     place.
    /// </remarks>
    /// <param name="document">The OpenAPI document to be transformed. Must not be null.</param>
    /// <param name="context">
    ///     The context information for the document transformation. Provides additional data or services required during
    ///     transformation.
    /// </param>
    /// <param name="cancellationToken">A cancellation token that can be used to cancel the asynchronous operation.</param>
    /// <returns>A task that represents the asynchronous transformation operation.</returns>
    public Task TransformAsync(OpenApiDocument document, OpenApiDocumentTransformerContext context, CancellationToken cancellationToken)
    {
        var paths = document.Paths;
        foreach (var path in paths)
            foreach (var operation in (path.Value.Operations ?? [])
                     .Where(operation => operation.Key == HttpMethod.Get))
            {
                if (operation.Value.Responses is null ||
                    !operation.Value.Responses.TryGetValue("200", out var response) ||
                    response.Content is null ||
                    !response.Content.TryGetValue("application/json", out var mediaType))
                {
                    continue;
                }

                mediaType.Schema = new OpenApiSchema
                {
                    Type = JsonSchemaType.Object,
                    ExternalDocs = new()
                    {
                        Description = "Find out more about Adaptive Cards",
                        Url = new("https://adaptivecards.io/schemas/adaptive-card.json")
                    }
                };
            }

        return Task.CompletedTask;
    }
}