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

    [Range(1, 64)]
    public int MaxConcurrency { get; set; } = 1;
    public Dictionary<string, string> Parameters { get; set; }
    public Dictionary<string, object?> ObjParameters { get; set; }
    public string LogSeparator { get; set; }
    public RSLogType RSLogType {  get; set; }
    public Action<Dictionary<string, object?>>? Init { get; set; }

    #endregion


    #region Initialization

    public HRConfiguration()
    {
      this.Scripts = new List<string>();
      this.Parameters = new Dictionary<string, string>();
      this.ObjParameters = new Dictionary<string, object?>();
      this.LogSeparator = String.Empty;

      this.RSLogType = RSLogType.Bulk;
      this.Init?.Invoke(this.ObjParameters);
    }

    #endregion


    #region Functions

    public Dictionary<string, object?> GetParameters()
    {
      var ret = new Dictionary<string, object?>();
      foreach (var parameter in this.Parameters)
      {
        ret.Add(parameter.Key, parameter.Value);
      }
      foreach (var parameter in this.ObjParameters)
      {
        ret.Add(parameter.Key, parameter.Value);
      }
      return ret;
    }

    #endregion

  }
}
