using System.Collections.ObjectModel;

namespace AiCalc.Models;

public class WorkbookSettings
{
    public ObservableCollection<WorkspaceConnection> Connections { get; set; } = new();

    public string DefaultImageModel { get; set; } = "stable-diffusion";

    public string DefaultTextModel { get; set; } = "gpt-4";

    public string DefaultVideoModel { get; set; } = "gen-2";
}
