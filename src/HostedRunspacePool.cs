using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;
using System.Threading;
using Microsoft.Extensions.Logging;

namespace DurHostedRunspace
{
  /// <summary>
  /// Contains functionality for executing PowerShell scripts.
  /// </summary>
  public class HostedRunspacePool : IDisposable
  {

    #region Fields

    private readonly ILogger? _logger;
    private readonly RunspacePool _rsPool;

    #endregion


    #region Properties

    /// <summary>
    /// Encoding used for script files
    /// </summary>
    public Encoding Encoding { get; set; }

    /// <summary>
    /// Determines if all logs are directly outputed or collected for each run
    /// </summary>
    public RSLogType RSLogType { get; set; }

    /// <summary>
    /// Determines if all logs are directly outputed or collected for each run
    /// </summary>
    public string LogSeparator { get; set; }

    /// <summary>
    /// Return the host app variable
    /// </summary>
    public string? DefaultHostApp { get; set; }

    #endregion


    #region Initialization

    public HostedRunspacePool(int minRunspaces = 1, int maxRunspaces = 5, 
      Dictionary<string, object?>? parameters = null!, IEnumerable<string>? modulesToLoad = null!, 
      ILogger? logger = null!, RSLogType rSLogType = RSLogType.Bulk, string logSeparator = "",
      string? hostApp = null!, Encoding encoding = null!, bool restricted = false)
    {
      this._logger = logger;
      this.LogSeparator = logSeparator;
      this.Encoding = encoding ?? Encoding.UTF8;

      this.DefaultHostApp = hostApp;
      this.RSLogType = rSLogType;

      var iss = HostedRunspace.GetInitialSessionState(restricted, parameters, modulesToLoad);


      this._rsPool = RunspaceFactory.CreateRunspacePool(iss);
      this._rsPool.SetMinRunspaces(minRunspaces);
      this._rsPool.SetMaxRunspaces(maxRunspaces);
      this._rsPool.ApartmentState = ApartmentState.MTA;
      this._rsPool.ThreadOptions = PSThreadOptions.ReuseThread;
      this._rsPool.Open();
    }



    public void Dispose()
    {
      this._rsPool.Dispose();
      GC.SuppressFinalize(this);
    }

    #endregion


    #region Functions

    public HostedRunspace GetHostedRunspace(ILogger? logger = null!,
      string? hostApp = null!, RSLogType? rSLogType = null,
      Encoding? encoding = null!, string logSeparator = null!)
    {
      var ps = PowerShell.Create();
      ps.RunspacePool = this._rsPool;

      return new HostedRunspace(ps, 
        logger ?? this._logger, rSLogType ?? this.RSLogType, logSeparator ?? this.LogSeparator,
        hostApp ?? this.DefaultHostApp, encoding ?? this.Encoding);
    }

    #endregion

  }
}
