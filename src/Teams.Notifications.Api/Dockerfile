#See https://aka.ms/customizecontainer to learn how to customize your debug container and how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
ARG BUILD_CONFIGURATION=Release
COPY . /src/
WORKDIR "/src/Teams.Notifications.Api"
RUN dotnet restore --force-evaluate --no-http-cache
RUN dotnet build -c $BUILD_CONFIGURATION -o /app/build --no-restore

FROM build AS publish
ARG BUILD_CONFIGURATION=Release
WORKDIR "/src/Teams.Notifications.Api"
RUN dotnet publish -c $BUILD_CONFIGURATION -o /app/publish /p:UseAppHost=false --no-restore

FROM mcr.microsoft.com/dotnet/aspnet:10.0 AS final
COPY --from=publish /app/publish /app
USER app
EXPOSE 8080
WORKDIR /app
ENTRYPOINT ["dotnet", "Teams.Notifications.Api.dll"]