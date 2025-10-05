# Installing and running AiCalc Studio

This guide walks through preparing your environment, restoring the project, and running the Uno Platform heads.

## Prerequisites

1. Install the [.NET 8 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/8.0).
2. Install the Uno project templates:
   ```bash
   dotnet new install Uno.ProjectTemplates.Dotnet
   ```
3. (Optional) Verify your workstation with the Uno check tool:
   ```bash
   dotnet tool install -g uno.check
   uno-check
   ```
4. For Windows builds, ensure the **Windows App SDK** and **Windows 10 SDK (19041 or newer)** are available. Visual Studio 2022 with the "Universal Windows Platform development" and ".NET desktop development" workloads is recommended.

## Cloning the repository

```bash
git clone https://github.com/<your-org>/aicalc.git
cd aicalc
```

## Restoring dependencies

```bash
dotnet restore
```

## Building the shared project

The solution is configured as a Uno Platform single-project application. Build the shared logic with:

```bash
dotnet build src/AiCalc/AiCalc.csproj -f net8.0
```

> This headless target is handy for CI environments and for validating non-Windows platforms that do not compile XAML locally.

## Running on Windows

```bash
dotnet build src/AiCalc/AiCalc.csproj -f net8.0-windows10.0.19041.0

dotnet run --project src/AiCalc/AiCalc.csproj -f net8.0-windows10.0.19041.0
```

You can also launch the **AiCalc** project from Visual Studio and select the `net8.0-windows10.0.19041.0` target.

## Running the WebAssembly head

WebAssembly support in Uno relies on the `wasm-tools` workload and a WebAssembly-specific head. The current repository focuses
on the shared logic and Windows target; to create a WebAssembly head:

1. Install the [wasm-tools workload](https://learn.microsoft.com/aspnet/core/blazor/webassembly-workload):
   ```bash
   dotnet workload install wasm-tools
   ```
2. Add the WebAssembly head following the [Uno Platform guidance](https://platform.uno/docs/articles/uno-platform-singleproject.html#adding-heads),
   ensuring it references `AiCalc.csproj` as the shared project.
3. Build and run the generated WebAssembly head (for example `AiCalc.Wasm`):
   ```bash
   dotnet run --project src/AiCalc.Wasm/AiCalc.Wasm.csproj -c Release
   ```
   The CLI prints the local development server URL; open it in a browser to interact with AiCalc Studio.

## Troubleshooting

- **Missing Windows SDK** – Install the Windows 10 SDK (19041 or later) through the Visual Studio installer or download it directly from Microsoft.
- **Workload restore failures** – Re-run `dotnet workload restore` or reinstall the affected workload (for example `dotnet workload install wasm-tools`).
- **XAML compilation errors on non-Windows OS** – Build with `-f net8.0` to skip Windows-only XAML, then deploy using Windows or WebAssembly builds from a Windows machine.

For more background, review the main [README](README.md).
