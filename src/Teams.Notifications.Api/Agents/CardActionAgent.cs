namespace Teams.Notifications.Api.Agents;

/// <summary>
///     Main agent that handles card actions and other interactions.
///     It does not create cards, that is done by controller actions that call the CardManagerService.
/// </summary>
public class CardActionAgent : AgentApplication
{
    private readonly ICardManagerService _cardManagerService;
    private readonly IFrontgateApiService _frontgateApiService;
    private readonly ILogger<CardActionAgent> _logger;
    private readonly ITeamsManagerService _teamsManagerService;
    private readonly ICustomEventTelemetryClient _telemetry;

    public CardActionAgent(AgentApplicationOptions options,
        ITeamsManagerService teamsManagerService,
        IFrontgateApiService frontgateApiService,
        ICardManagerService cardManagerService,
        ICustomEventTelemetryClient telemetry,
        ILogger<CardActionAgent> logger
    ) : base(options)
    {
        _logger = logger;
        _telemetry = telemetry;
        _teamsManagerService = teamsManagerService;
        _frontgateApiService = frontgateApiService;
        _cardManagerService = cardManagerService;
        OnMessageReactionsAdded(MessageReactionAsync);
        AdaptiveCards.OnActionExecute("Process", ProcessCardActionAsync);
        AdaptiveCards.OnActionExecute("WelcomeBack", WelcomeBackCardActionAsync);

        OnMessage(ConversationUpdateEvents.MembersAdded, WelcomeMessageToUserAsync);
        OnActivity(ActivityTypes.Message, MessageActivityAsync, RouteRank.Last);
    }


    // Proof of concept for message reactions
    private static async Task MessageReactionAsync(ITurnContext turnContext, ITurnState turnState, CancellationToken cancellationToken)
    {
        await turnContext.SendActivityAsync("Message Reaction: " + turnContext.Activity.ReactionsAdded[0].Type, cancellationToken: cancellationToken);
    }


    // proof of concept for welcome message, for now just says welcome
    private static async Task WelcomeMessageToUserAsync(ITurnContext turnContext, ITurnState turnState, CancellationToken cancellationToken)
    {
        var member = await TeamsInfo.GetMemberAsync(turnContext, turnContext.Activity.From.Id, cancellationToken);
        var user = member.Name ?? "new user";
        await turnContext.SendActivityAsync(MessageFactory.Text("Welcome " + user), cancellationToken);
    }


    // If the User tries to send us a message, we just tell them we don't support it
    private static async Task MessageActivityAsync(ITurnContext turnContext, ITurnState turnState, CancellationToken cancellationToken)
    {
        if (!string.IsNullOrWhiteSpace(turnContext.Activity.Text)) await turnContext.SendActivityAsync(MessageFactory.Text("We don't support any interaction at the moment"), cancellationToken);
    }


    //     LogicApp handle of the "Reprocess File" button, will send it to Frontgate for reprocessing, and update the card
    //     accordingly, so you can't press it again
    private Task<AdaptiveCardInvokeResponse> ProcessCardActionAsync(ITurnContext turnContext, ITurnState turnState, object data, CancellationToken token) => turnContext.HandleProcessVerbLogicAppAsync(data, _telemetry, _logger, _teamsManagerService, _frontgateApiService, _cardManagerService, token);

    // WelcomeCard.json "Welcome Back" button action handler
    private async Task<AdaptiveCardInvokeResponse> WelcomeBackCardActionAsync(ITurnContext turnContext, ITurnState turnState, object data, CancellationToken token)
    {
        await turnContext.SendActivityAsync(MessageFactory.Text("I am going to work on your request"), token);
        return AdaptiveCardInvokeResponseFactory.Message("Thank you for your reaction");
    }
}