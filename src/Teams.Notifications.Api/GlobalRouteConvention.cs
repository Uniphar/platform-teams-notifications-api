namespace Teams.Notifications.Api;

internal sealed class GlobalRouteConvention(string routePrefix) : IApplicationModelConvention
{
    public void Apply(ApplicationModel application)
    {
        var prefixSelector = AttributeRouteModel.CombineTemplates(routePrefix, string.Empty);

        foreach (var controller in application.Controllers)
            foreach (var selector in controller.Selectors)
            {
                selector.AttributeRouteModel = selector.AttributeRouteModel is null
                    ? new() { Template = prefixSelector }
                    : AttributeRouteModel.CombineAttributeRouteModel(
                        new() { Template = routePrefix },
                        selector.AttributeRouteModel);
            }
    }
}