using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Reflection;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace Flow_Finder
{
    // Do not forget to update version number and author (company attribute) in AssemblyInfo.cs class
    // To generate Base64 string for Images below, you can use https://www.base64-image.de/
    [Export(typeof(IXrmToolBoxPlugin)),
        ExportMetadata("Name", "Flow Finder"),
        ExportMetadata("Version", "1.0.0.0"),
        ExportMetadata("Author", "Matt Collins-Jones"),
        ExportMetadata("Company", "Matt Collins-Jones"),
        ExportMetadata("Description", "Find Flows in Solutions, export them, add or remove solutions and Co-owners"),
        // Please specify the base64 content of a 32x32 pixels image
        ExportMetadata("SmallImageBase64", "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"),
        // Please specify the base64 content of a 80x80 pixels image
        ExportMetadata("BigImageBase64", "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"),
        ExportMetadata("BackgroundColor", "Lavender"),
        ExportMetadata("PrimaryFontColor", "Black"),
        ExportMetadata("SecondaryFontColor", "Gray"),
        ExportMetadata("Id", "d8b5b9f2-6c3a-4e91-9f2b-1a2b3c4d5e6f")]
    public class FlowFinder : PluginBase
    {
        public override IXrmToolBoxPluginControl GetControl()
        {
            return new FlowFinderControl();
        }

        /// <summary>
        /// Constructor 
        /// </summary>
        public FlowFinder()
        {
            // If you have external assemblies that you need to load, uncomment the following to 
            // hook into the event that will fire when an Assembly fails to resolve
            // AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(AssemblyResolveEventHandler);
        }

        /// <summary>
        /// Event fired by CLR when an assembly reference fails to load
        /// Assumes that related assemblies will be loaded from a subfolder named the same as the Plugin
        /// For example, a folder named Sample.XrmToolBox.MyPlugin 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        private Assembly AssemblyResolveEventHandler(object sender, ResolveEventArgs args)
        {
            Assembly loadAssembly = null;
            Assembly currAssembly = Assembly.GetExecutingAssembly();

            // base name of the assembly that failed to resolve
            var argName = args.Name.Substring(0, args.Name.IndexOf(","));

            // check to see if the failing assembly is one that we reference.
            List<AssemblyName> refAssemblies = currAssembly.GetReferencedAssemblies().ToList();
            var refAssembly = refAssemblies.Where(a => a.Name == argName).FirstOrDefault();

            // if the current unresolved assembly is referenced by our plugin, attempt to load
            if (refAssembly != null)
            {
                // load from the path to this plugin assembly, not host executable
                string dir = Path.GetDirectoryName(currAssembly.Location).ToLower();
                string folder = Path.GetFileNameWithoutExtension(currAssembly.Location);
                dir = Path.Combine(dir, folder);

                var assmbPath = Path.Combine(dir, $"{argName}.dll");

                if (File.Exists(assmbPath))
                {
                    loadAssembly = Assembly.LoadFrom(assmbPath);
                }
                else
                {
                    throw new FileNotFoundException($"Unable to locate dependency: {assmbPath}");
                }
            }

            return loadAssembly;
        }
    }
}