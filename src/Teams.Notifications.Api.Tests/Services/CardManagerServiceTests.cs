using System.Security.Claims;
using Microsoft.Extensions.Logging;

namespace Teams.Notifications.Api.Tests.Services;

[TestClass]
[TestCategory("Unit")]
public class CardManagerServiceTests
{
    private readonly Mock<IChannelAdapter> _adapterMock;
    private readonly Mock<IConfiguration> _configMock;
    private readonly Mock<ILogger<CardManagerService>> _loggerMock;
    private readonly Mock<ICosmosMessageStore> _messageStoreMock;
    private readonly Mock<ITeamsManagerService> _teamsManagerServiceMock;
    private readonly Mock<ICustomEventTelemetryClient> _telemetryMock;

    public CardManagerServiceTests()
    {
        _adapterMock = new();
        _teamsManagerServiceMock = new();
        _messageStoreMock = new();
        _configMock = new();
        _configMock.Setup(c => c["AZURE_CLIENT_ID"]).Returns("client-id");
        _configMock.Setup(c => c["AZURE_TENANT_ID"]).Returns("tenant-id");
        _telemetryMock = new();
        _loggerMock = new();
    }

    private CardManagerService CreateService() => new(_adapterMock.Object, _teamsManagerServiceMock.Object, _messageStoreMock.Object, _configMock.Object, _loggerMock.Object, _telemetryMock.Object);

    private static StoredMessage ChannelDoc(string teamId, string channelId, string jsonFileName, string uniqueId, string messageId) =>
        new()
        {
            Id = messageId,
            PartitionKey = StoredMessage.ChannelPartition(teamId, channelId),
            TeamId = teamId,
            ChannelId = channelId,
            JsonFileName = jsonFileName,
            UniqueId = uniqueId,
            CardJson = "{}"
        };

    [TestMethod]
    public async Task DeleteCard_DeletesCard_WhenIdIsFound()
    {
        // Arrange
        var service = CreateService();
        _teamsManagerServiceMock.Setup(x => x.GetTeamIdAsync("team", CancellationToken.None)).ReturnsAsync("teamId");
        _teamsManagerServiceMock.Setup(x => x.CheckOrInstallBotIsInTeam("teamId", CancellationToken.None)).Returns(Task.CompletedTask);
        _teamsManagerServiceMock.Setup(x => x.GetChannelIdAsync("teamId", "channel", CancellationToken.None)).ReturnsAsync("channelId");
        _messageStoreMock
            .Setup(x => x.FindByChannelAsync("teamId", "channelId", "file.json", "uid", CancellationToken.None))
            .ReturnsAsync(ChannelDoc("teamId", "channelId", "file.json", "uid", "msgId"));
        _adapterMock
            .Setup(x => x.ContinueConversationAsync(
                It.IsAny<ClaimsIdentity>(),
                It.IsAny<ConversationReference>(),
                It.IsAny<AgentCallbackHandler>(),
                It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        // Act
        await service.DeleteCardAsync("file.json", "uid", "team", "channel", CancellationToken.None);

        // Assert
        _adapterMock.Verify(x => x.ContinueConversationAsync(
                It.IsAny<ClaimsIdentity>(),
                It.Is<ConversationReference>(cr => cr.ActivityId == "msgId"),
                It.IsAny<AgentCallbackHandler>(),
                CancellationToken.None),
            Times.Once);
    }

    [TestMethod]
    public async Task DeleteCard_Throws_WhenIdNotFound()
    {
        // Arrange
        var service = CreateService();
        _teamsManagerServiceMock.Setup(x => x.GetTeamIdAsync(It.IsAny<string>(), It.IsAny<CancellationToken>())).ReturnsAsync("teamId");
        _teamsManagerServiceMock.Setup(x => x.CheckOrInstallBotIsInTeam(It.IsAny<string>(), It.IsAny<CancellationToken>())).Returns(Task.CompletedTask);
        _teamsManagerServiceMock.Setup(x => x.GetChannelIdAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>())).ReturnsAsync("channelId");
        _messageStoreMock
            .Setup(x => x.FindByChannelAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((StoredMessage?)null);

        // Act & Assert
        await Assert.ThrowsExactlyAsync<InvalidOperationException>(() => service.DeleteCardAsync("file.json", "uid", "team", "channel", CancellationToken.None));
    }

    [TestMethod]
    public async Task CreateOrUpdate_CreatesNewCard_WhenNoExistingId()
    {
        // Arrange
        var service = CreateService();
        var model = new BaseTemplateModel { UniqueId = "uid" };
        _teamsManagerServiceMock.Setup(x => x.GetTeamIdAsync("team", CancellationToken.None)).ReturnsAsync("teamId");
        _teamsManagerServiceMock.Setup(x => x.CheckOrInstallBotIsInTeam("teamId", CancellationToken.None)).Returns(Task.CompletedTask);
        _teamsManagerServiceMock.Setup(x => x.GetChannelIdAsync("teamId", "channel", CancellationToken.None)).ReturnsAsync("channelId");
        _messageStoreMock
            .Setup(x => x.FindByChannelAsync("teamId", "channelId", "file.json", "uid", CancellationToken.None))
            .ReturnsAsync((StoredMessage?)null);

        _adapterMock
            .Setup(x => x.ContinueConversationAsync(
                It.IsAny<ClaimsIdentity>(),
                It.IsAny<ConversationReference>(),
                It.IsAny<AgentCallbackHandler>(),
                It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        // Act
        await service.CreateOrUpdateAsync("WelcomeCard.json", null, model, "team", "channel", CancellationToken.None);

        // Assert
        _adapterMock.Verify(x => x.ContinueConversationAsync(
                It.IsAny<ClaimsIdentity>(),
                It.IsAny<ConversationReference>(),
                It.IsAny<AgentCallbackHandler>(),
                CancellationToken.None),
            Times.Once);
    }

    [TestMethod]
    public void GetConversationReference_ReturnsExpectedReference()
    {
        // Arrange
        var service = CreateService();
        var channelId = "channelId";

        // Act
        var result = typeof(CardManagerService)
            .GetMethod("GetConversationReference", BindingFlags.NonPublic | BindingFlags.Instance)!
            .Invoke(service, new object[] { channelId }) as ConversationReference;

        // Assert
        Assert.IsNotNull(result);
        Assert.AreEqual("msteams", result.ChannelId);
        Assert.AreEqual("https://smba.trafficmanager.net/emea/tenant-id", result.ServiceUrl);
        Assert.AreEqual(channelId, result.Conversation.Id);
        Assert.AreEqual(channelId, result.ActivityId);
    }

    [TestMethod]
    public async Task BasicCreateCardFromTemplateAsyncTest()
    {
        var model = new LogicAppErrorModel
        {
            TimeStamp = "01-01-1960",
            ObjectType = "test",
            ErrorMessage = "This request is not authorized to perform this operation using this permission.\\nRequestId:e4669c49-a002-0049-449a-17061a000000\\nTime:2025-08-27T21:34:16.5796961Z\",\"This request is not authorized to perform this operation using this permission.\\nRequestId:77c89d5a-9002-0042-359a-17fd71000000\\nTime:2025-08-27T21:34:18.5359065Z",
            UniqueId = "unique"
        };
        // Arrange
        var service = CreateService();
        var result = await service.CreateCardFromTemplateAsync("LogicAppError.json", null, model, string.Empty, string.Empty, string.Empty, CancellationToken.None);
        // Assert
        Assert.IsNotEmpty(result);
        var item = AdaptiveCard.FromJson(result).Card;
        Assert.IsNotNull(item.Body);
        // 5 items should be left since the rest should be removed
        Assert.HasCount(5, item.Body);
        foreach (var element in item.Body)
        {
            switch (element)
            {
                case AdaptiveTextBlock textBlock:
                    Assert.DoesNotContain("{{", textBlock.Text, "No template string should be found!, found: {0}", textBlock.Text);
                    Assert.DoesNotContain("}}", textBlock.Text, "No template string should be found!, found: {0}", textBlock.Text);
                    break;
                case AdaptiveFactSet adaptiveSet:
                {
                    foreach (var fact in adaptiveSet.Facts)
                    {
                        Assert.DoesNotContain("{{", fact.Value, "No template string should be found!, found: {0}", fact.Value);
                        Assert.DoesNotContain("}}", fact.Value, "No template string should be found!, found: {0}", fact.Value);
                    }

                    break;
                }
            }
        }
    }
}