using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DurHostedRunspace
{
  public class HostedRunspaceService
  {

    #region Properties

    private readonly HRConfiguration _settings;
    private readonly List<string> _modules;
    
    private HostedRunspace? _hostedRunspace;
    private readonly HostedRunspacePool? _hostedRunspacePool;

    public bool IsPool => this._settings.MaxConcurrency > 1;

    #endregion


    #region Initialization

    public HostedRunspaceService(ILogger<HostedRunspaceService> logger,
       IOptions<HRConfiguration> settings)
    {
      this._settings = settings.Value;
      string? hostApp = HostedRunspace.GetDefaultHostApp();

      this._modules ??= new List<string>();

      foreach(var scriptPath in this._settings.Scripts)
      {
        if (System.IO.File.Exists(scriptPath))
        {
          this._modules.Add(scriptPath);
        } else if (System.IO.Directory.Exists(scriptPath))
        {
          this._modules.AddRange(System.IO.Directory.GetFiles(scriptPath,
          "*.ps1", System.IO.SearchOption.TopDirectoryOnly));
        }
      }

      if (this._modules.Count > 0)
      {
        this._modules = this._modules.Distinct().ToList();
        var paketeVersion = System.Reflection.Assembly.GetEntryAssembly()!
          .GetName()!.Version!;
        var log = new System.Text.StringBuilder();
        log.Append("Initialize HostedRunspace:");
        log.AppendLine($"DurApps: {hostApp}/{paketeVersion}");
        log.AppendLine($".Net Version: {Environment.Version}");
        log.AppendLine();
        log.AppendLine("Loaded modules:");
        log.AppendLine(String.Join(Environment.NewLine, this._modules.Select(m => "-> " + m)));
        logger.LogInformation(log.ToString().Trim());
      }

      System.Threading.Thread.Sleep(5000);

      if (!this.IsPool)
      {
        this._hostedRunspace = new HostedRunspace(
          this._settings.Parameters, this._modules, 
          logger, this._settings.RSLogType, this._settings.LogSeparator,
          hostApp, System.Text.Encoding.UTF8, false);
      }
      else
      {
        this._hostedRunspacePool = new HostedRunspacePool(1, this._settings.MaxConcurrency,
          this._settings.Parameters, this._modules,
          logger, this._settings.RSLogType, this._settings.LogSeparator,
          hostApp, System.Text.Encoding.UTF8, false);
      }
    }

    #endregion


    #region Functions

    public HostedRunspace GetHostedRunspace(ILogger? logger = null)
    {
      if (this.IsPool && this._hostedRunspacePool != null)
      {
        return this._hostedRunspacePool.GetHostedRunspace(logger);
      }
      else if (!this.IsPool)
      {
        if (this._hostedRunspace == null || this._hostedRunspace.IsDisposed)
        {
          this._hostedRunspace = new HostedRunspace(
            this._settings.Parameters, this._modules,
            logger, this._settings.RSLogType, this._settings.LogSeparator,
            null, System.Text.Encoding.UTF8, false);
        }

        return this._hostedRunspace;
      }
      throw new ApplicationException();
    }


    public async Task<HostedRunspace> GetHostedRunspaceAsync(ILogger? logger = null)
    {
      if (this.IsPool && this._hostedRunspacePool != null)
      {
        return this._hostedRunspacePool.GetHostedRunspace(logger);
      }
      else if (!this.IsPool)
      {
        if (this._hostedRunspace == null || this._hostedRunspace.IsDisposed)
        {
          this._hostedRunspace = await Task.Run(() => new HostedRunspace(
            this._settings.Parameters, this._modules,
            logger, this._settings.RSLogType, this._settings.LogSeparator,
            null, System.Text.Encoding.UTF8, false));
        }

        return this._hostedRunspace;
      }
      throw new ApplicationException();
    }

    #endregion

  }
}
