# Etapa de construcci�n
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Copiar archivos de soluci�n y restaurar dependencias
COPY ["LatExcel.sln", "./"]
COPY ["LatExcel.Api/LatExcel.Api.csproj", "LatExcel.Api/"]
COPY ["LatExcel.Aplicacion/LatExcel.Aplicacion.csproj", "LatExcel.Aplicacion/"]
RUN dotnet restore "LatExcel.sln"

# Copiar el resto de los archivos y compilar
COPY . .
WORKDIR "/src/LatExcel.Api"
RUN dotnet publish "LatExcel.Api.csproj" -c Release -o /app/publish

# Etapa de ejecuci�n
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS runtime
WORKDIR /app
COPY --from=build /app/publish .
ENTRYPOINT ["dotnet", "LatExcel.Api.dll"]
