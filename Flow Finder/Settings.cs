using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Flow_Finder
{
    /// <summary>
    /// This class can help you to store settings for your plugin
    /// </summary>
    /// <remarks>
    /// This class must be XML serializable
    /// </remarks>
    public class Settings
    {
        public string LastUsedOrganizationWebappUrl { get; set; }
        public RefreshMode RefreshAfterDialogMode { get; set; } = RefreshMode.RefreshFlow;
    }

    public enum RefreshMode
    {
        None = 0,
        RefreshFlow = 1,
        RefreshAll = 2
    }
}