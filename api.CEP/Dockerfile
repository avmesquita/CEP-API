FROM microsoft/dotnet:2.2-aspnetcore-runtime AS base
WORKDIR /app
EXPOSE 80
EXPOSE 443

FROM microsoft/dotnet:2.2-sdk AS build
WORKDIR /src
COPY ["api.CEP/api.CEP.csproj", "api.CEP/"]
RUN dotnet restore "api.CEP/api.CEP.csproj"
COPY . .
WORKDIR "/src/api.CEP"
RUN dotnet build "api.CEP.csproj" -c Release -o /app

FROM build AS publish
RUN dotnet publish "api.CEP.csproj" -c Release -o /app

FROM base AS final
WORKDIR /app
COPY --from=publish /app .
ENTRYPOINT ["dotnet", "api.CEP.dll"]