using Nice3point.Revit.Toolkit.External;
using SpotDimensionPlugin.Commands;

namespace SpotDimensionPlugin;

/// <summary>
///     Application entry point
/// </summary>
[UsedImplicitly]
public class Application : ExternalApplication
{
    public override void OnStartup()
    {
        CreateRibbon();
    }

    private void CreateRibbon()
    {
        var panel = Application.CreatePanel("Commands", "SpotDimensionPlugin");

        panel.AddPushButton<StartupCommand>("Execute")
            .SetImage("/SpotDimensionPlugin;component/Resources/Icons/RibbonIcon16.png")
            .SetLargeImage("/SpotDimensionPlugin;component/Resources/Icons/RibbonIcon32.png");
    }
}