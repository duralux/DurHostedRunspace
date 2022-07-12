using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DurHostedRunspace
{
  public class HRConfiguration
  {

    public const string NAME = "HostedRunspace";

    #region Properties

    public List<string> Scripts { get; set; }
    [Obsolete]
    public string? CommonScripts { get; set; }
    [Obsolete]
    public string? UserScripts { get; set; }

    [Range(1, 64)]
    public int MaxConcurrency { get; set; } = 1;
    public Dictionary<string, string> Parameters { get; set; }
    public string LogSeparator { get; set; }
    public RSLogType RSLogType {  get; set; }

    #endregion


    #region Initialization

    public HRConfiguration()
    {
      this.Scripts = new List<string>();
      this.Parameters = new Dictionary<string, string>();
      this.LogSeparator = String.Empty;

      this.RSLogType = RSLogType.Bulk;
    }

    #endregion

  }
}
