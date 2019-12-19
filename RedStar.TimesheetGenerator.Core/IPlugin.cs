namespace RedStar.TimesheetGenerator.Core
{
    public interface IPlugin
    {
        /// <summary>
        /// The name of the plugin, to be used (case-insensitive)
        /// at the command line.
        /// </summary>
        string Name { get; }
    }
}