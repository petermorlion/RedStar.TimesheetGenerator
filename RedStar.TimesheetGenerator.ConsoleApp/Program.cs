using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using RedStar.TimesheetGenerator.Core;

namespace RedStar.TimesheetGenerator.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var sourceName = args[0];
                var destinationName = args[1];
                var dateArg = args[2];
                var destination = args[3];

                var options = new Options
                {
                    FileDestination = new FileInfo(destination),
                    Year = int.Parse(dateArg.Substring(0, 4)),
                    Month = int.Parse(dateArg.Substring(4, 2))
                };

                var currentPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                var possiblePluginPaths = Directory.EnumerateFiles(currentPath, "*.dll", SearchOption.AllDirectories);

                var possiblePluginAssemblies = possiblePluginPaths.Select(LoadPlugin);

                var source = GetPlugin<ITimeTrackingSource>(possiblePluginAssemblies, sourceName, options);

                if (source == null)
                {
                    Console.WriteLine($"No source found with name '{sourceName}'.");
                    Environment.Exit(-1);
                }

                var destinationPlugin = GetPlugin<ITimesheetDestination>(possiblePluginAssemblies, destinationName, options);

                if (destinationPlugin == null)
                {
                    Console.WriteLine($"No destination found with name '{destinationName}'.");
                    Environment.Exit(-1);
                }

                var entries = source.GetEntries();

                destinationPlugin.CreateTimesheet(entries);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        private static T GetPlugin<T>(IEnumerable<Assembly> possiblePluginAssemblies, string pluginName, Options options) where T : IPlugin
        {
            var result = possiblePluginAssemblies
                .SelectMany(assembly => CreatePlugin<T>(assembly, options))
                .FirstOrDefault(x => x.Name.ToLower() == pluginName.ToLower());

            return result;
        }

        static Assembly LoadPlugin(string pluginLocation)
        {
            Console.WriteLine($"Loading DLL: {pluginLocation}");
            var loadContext = new PluginLoadContext(pluginLocation);
            return loadContext.LoadFromAssemblyName(new AssemblyName(Path.GetFileNameWithoutExtension(pluginLocation)));
        }

        static IEnumerable<T> CreatePlugin<T>(Assembly assembly, Options options) where T : IPlugin
        {
            foreach (var type in assembly.GetTypes())
            {
                if (typeof(T).IsAssignableFrom(type))
                {
                    Console.WriteLine($"Found plugin {type.Name}");
                    if (Activator.CreateInstance(type, options) is T result)
                    {
                        yield return result;
                    }
                }
             }
        }
    }
}
