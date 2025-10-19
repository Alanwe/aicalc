using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace AiCalc.Services;

/// <summary>
/// Detects installed Python environments (Phase 7)
/// </summary>
public class PythonEnvironmentDetector
{
    public class PythonEnvironment
    {
        public string Name { get; set; } = string.Empty;
        public string Path { get; set; } = string.Empty;
        public string Version { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty; // System, Conda, Venv
        public bool IsValid { get; set; }
    }

    /// <summary>
    /// Detect all available Python environments
    /// </summary>
    public static List<PythonEnvironment> DetectEnvironments()
    {
        var environments = new List<PythonEnvironment>();

        // 1. Check registry (Windows Python installations)
        environments.AddRange(DetectFromRegistry());

        // 2. Check PATH environment variable
        environments.AddRange(DetectFromPath());

        // 3. Check common Conda locations
        environments.AddRange(DetectCondaEnvironments());

        // 4. Check common virtual environment locations
        environments.AddRange(DetectVirtualEnvironments());

        // Remove duplicates based on path
        var uniqueEnvironments = environments
            .GroupBy(e => e.Path.ToLowerInvariant())
            .Select(g => g.First())
            .OrderBy(e => e.Type)
            .ThenBy(e => e.Name)
            .ToList();

        return uniqueEnvironments;
    }

    private static List<PythonEnvironment> DetectFromRegistry()
    {
        var environments = new List<PythonEnvironment>();

        try
        {
            // Check both 32-bit and 64-bit registry locations
            var registryPaths = new[]
            {
                @"SOFTWARE\Python\PythonCore",
                @"SOFTWARE\Wow6432Node\Python\PythonCore"
            };

            foreach (var registryPath in registryPaths)
            {
                using var key = Registry.LocalMachine.OpenSubKey(registryPath);
                if (key == null) continue;

                foreach (var versionName in key.GetSubKeyNames())
                {
                    using var versionKey = key.OpenSubKey(versionName);
                    if (versionKey == null) continue;

                    using var installPathKey = versionKey.OpenSubKey("InstallPath");
                    if (installPathKey == null) continue;

                    var installPath = installPathKey.GetValue("") as string;
                    if (string.IsNullOrEmpty(installPath)) continue;

                    var pythonExe = Path.Combine(installPath, "python.exe");
                    if (File.Exists(pythonExe))
                    {
                        var env = CreateEnvironment(pythonExe, $"Python {versionName}", "System");
                        if (env.IsValid)
                        {
                            environments.Add(env);
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error detecting from registry: {ex.Message}");
        }

        return environments;
    }

    private static List<PythonEnvironment> DetectFromPath()
    {
        var environments = new List<PythonEnvironment>();

        try
        {
            var pathEnv = Environment.GetEnvironmentVariable("PATH");
            if (string.IsNullOrEmpty(pathEnv)) return environments;

            var paths = pathEnv.Split(';');
            foreach (var path in paths)
            {
                if (string.IsNullOrWhiteSpace(path)) continue;

                var pythonExe = Path.Combine(path.Trim(), "python.exe");
                if (File.Exists(pythonExe))
                {
                    var env = CreateEnvironment(pythonExe, "Python (PATH)", "System");
                    if (env.IsValid)
                    {
                        environments.Add(env);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error detecting from PATH: {ex.Message}");
        }

        return environments;
    }

    private static List<PythonEnvironment> DetectCondaEnvironments()
    {
        var environments = new List<PythonEnvironment>();

        try
        {
            // Check common Conda locations
            var condaLocations = new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "miniconda3"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "anaconda3"),
                @"C:\ProgramData\miniconda3",
                @"C:\ProgramData\anaconda3"
            };

            foreach (var condaRoot in condaLocations)
            {
                if (!Directory.Exists(condaRoot)) continue;

                // Base conda environment
                var condaPython = Path.Combine(condaRoot, "python.exe");
                if (File.Exists(condaPython))
                {
                    var env = CreateEnvironment(condaPython, "Conda (base)", "Conda");
                    if (env.IsValid)
                    {
                        environments.Add(env);
                    }
                }

                // Conda environments
                var envsDir = Path.Combine(condaRoot, "envs");
                if (Directory.Exists(envsDir))
                {
                    foreach (var envDir in Directory.GetDirectories(envsDir))
                    {
                        var envPython = Path.Combine(envDir, "python.exe");
                        if (File.Exists(envPython))
                        {
                            var envName = Path.GetFileName(envDir);
                            var env = CreateEnvironment(envPython, $"Conda ({envName})", "Conda");
                            if (env.IsValid)
                            {
                                environments.Add(env);
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error detecting Conda environments: {ex.Message}");
        }

        return environments;
    }

    private static List<PythonEnvironment> DetectVirtualEnvironments()
    {
        var environments = new List<PythonEnvironment>();

        try
        {
            // Check common venv locations in current directory and subdirectories
            var searchPaths = new[]
            {
                Environment.CurrentDirectory,
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Projects")
            };

            foreach (var searchPath in searchPaths)
            {
                if (!Directory.Exists(searchPath)) continue;

                // Look for venv directories (limited depth to avoid long searches)
                var venvDirs = new[] { "venv", ".venv", "env", ".env" };
                
                foreach (var venvDir in venvDirs)
                {
                    var venvPath = Path.Combine(searchPath, venvDir);
                    if (Directory.Exists(venvPath))
                    {
                        var venvPython = Path.Combine(venvPath, "Scripts", "python.exe");
                        if (File.Exists(venvPython))
                        {
                            var env = CreateEnvironment(venvPython, $"Venv ({venvDir})", "Venv");
                            if (env.IsValid)
                            {
                                environments.Add(env);
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error detecting virtual environments: {ex.Message}");
        }

        return environments;
    }

    private static PythonEnvironment CreateEnvironment(string pythonExePath, string name, string type)
    {
        var env = new PythonEnvironment
        {
            Name = name,
            Path = pythonExePath,
            Type = type,
            IsValid = false
        };

        try
        {
            // Get Python version
            var startInfo = new ProcessStartInfo
            {
                FileName = pythonExePath,
                Arguments = "--version",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = Process.Start(startInfo);
            if (process != null)
            {
                process.WaitForExit(5000); // 5 second timeout

                var output = process.StandardOutput.ReadToEnd();
                var error = process.StandardError.ReadToEnd();
                var versionText = !string.IsNullOrEmpty(output) ? output : error;

                // Extract version number (e.g., "Python 3.11.4")
                var match = Regex.Match(versionText, @"Python\s+(\d+\.\d+\.\d+)");
                if (match.Success)
                {
                    env.Version = match.Groups[1].Value;
                    env.IsValid = true;
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error getting Python version for {pythonExePath}: {ex.Message}");
        }

        return env;
    }

    /// <summary>
    /// Test if a Python environment has the aicalc-sdk package installed
    /// </summary>
    public static bool HasAiCalcSdk(string pythonExePath)
    {
        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = pythonExePath,
                Arguments = "-m pip show aicalc-sdk",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = Process.Start(startInfo);
            if (process != null)
            {
                process.WaitForExit(5000);
                var output = process.StandardOutput.ReadToEnd();
                return output.Contains("Name: aicalc-sdk");
            }
        }
        catch
        {
            // Ignore errors
        }

        return false;
    }

    /// <summary>
    /// Install aicalc-sdk in the specified Python environment
    /// </summary>
    public static (bool Success, string Message) InstallAiCalcSdk(string pythonExePath)
    {
        try
        {
            // Get the path to the SDK
            var sdkPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "python-sdk");
            sdkPath = Path.GetFullPath(sdkPath);

            if (!Directory.Exists(sdkPath))
            {
                return (false, "Python SDK directory not found");
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = pythonExePath,
                Arguments = $"-m pip install -e \"{sdkPath}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = Process.Start(startInfo);
            if (process != null)
            {
                process.WaitForExit(30000); // 30 second timeout

                var output = process.StandardOutput.ReadToEnd();
                var error = process.StandardError.ReadToEnd();

                if (process.ExitCode == 0)
                {
                    return (true, "Successfully installed aicalc-sdk");
                }
                else
                {
                    return (false, $"Installation failed: {error}");
                }
            }
        }
        catch (Exception ex)
        {
            return (false, $"Error installing SDK: {ex.Message}");
        }

        return (false, "Unknown error");
    }
}
