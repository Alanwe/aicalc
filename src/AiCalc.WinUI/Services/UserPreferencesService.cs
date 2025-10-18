using System;
using System.IO;
using System.Text.Json;
using System.Linq;
using AiCalc.Models;

namespace AiCalc.Services;

/// <summary>
/// Service for loading and saving user preferences (Phase 5)
/// </summary>
public class UserPreferencesService
{
    private readonly string _preferencesPath;
    private UserPreferences? _currentPreferences;

    public UserPreferencesService()
    {
        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var appFolder = Path.Combine(localAppData, "AiCalc");
        Directory.CreateDirectory(appFolder);
        _preferencesPath = Path.Combine(appFolder, "preferences.json");
    }

    /// <summary>
    /// Load user preferences from disk, or create defaults if not found
    /// </summary>
    public UserPreferences LoadPreferences()
    {
        if (_currentPreferences != null)
        {
            return _currentPreferences;
        }

        try
        {
            if (File.Exists(_preferencesPath))
            {
                var json = File.ReadAllText(_preferencesPath);
                _currentPreferences = JsonSerializer.Deserialize<UserPreferences>(json) ?? new UserPreferences();
            }
            else
            {
                _currentPreferences = new UserPreferences();
            }
        }
        catch (Exception ex)
        {
            // If preferences file is corrupted, use defaults
            System.Diagnostics.Debug.WriteLine($"Error loading preferences: {ex.Message}");
            _currentPreferences = new UserPreferences();
        }

        return _currentPreferences;
    }

    /// <summary>
    /// Save user preferences to disk
    /// </summary>
    public void SavePreferences(UserPreferences preferences)
    {
        try
        {
            _currentPreferences = preferences;
            var options = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(preferences, options);
            File.WriteAllText(_preferencesPath, json);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error saving preferences: {ex.Message}");
        }
    }

    /// <summary>
    /// Add a workbook path to recent files list (max 10)
    /// </summary>
    public void AddRecentWorkbook(string path)
    {
        var prefs = LoadPreferences();
        
        // Remove if already in list
        var recentList = prefs.RecentWorkbooks.Where(p => !string.Equals(p, path, StringComparison.OrdinalIgnoreCase)).ToList();
        
        // Add to front
        recentList.Insert(0, path);
        
        // Keep only 10 most recent
        prefs.RecentWorkbooks = recentList.Take(10).ToArray();
        prefs.LastWorkbookPath = path;
        
        SavePreferences(prefs);
    }
}
