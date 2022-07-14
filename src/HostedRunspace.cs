using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace DurHostedRunspace
{
  /// <summary>
  /// Contains functionality for executing PowerShell scripts.
  /// </summary>
  public class HostedRunspace : IDisposable
  {

    #region Fields

    private const string HOST_APP = @"HostApp";

    private readonly ILogger? _logger;
    private readonly PowerShell _ps;

    #endregion


    #region Properties

    /// <summary>
    /// Returns the version of the runspace
    /// </summary>
    public Version RunspaceVersion => this._ps.Runspace.Version;

    /// <summary>
    /// Returns true, if the runspace is restricted, otherwise false
    /// </summary>
    public bool IsRestricted
    {
      get
      {
        try
        {
          return ((Microsoft.PowerShell.ExecutionPolicy?)
            RunScript("Get-ExecutionPolicy")?.SingleOrDefault() ??
            Microsoft.PowerShell.ExecutionPolicy.Undefined) ==
            Microsoft.PowerShell.ExecutionPolicy.Restricted;
        }
        catch (Exception ex)
        {
          this._logger?.LogError(ex,
            "IsRestricted? Message pipeline contained not exactly one element!");
          return false;
        }
      }
    }


    /// <summary>
    /// Encoding used for script files
    /// </summary>
    public Encoding Encoding { get; set; }

    /// <summary>
    /// Determines if all logs are directly outputed or collected for each run
    /// </summary>
    public RSLogType RSLogType { get; set; }
    public LogLevel MaxLogLevel { get; private set; }
    public string LogSeparator { get; set; }
    public StringBuilder? Log { get; private set; }

    /// <summary>
    /// Return the host app variable
    /// </summary>
    public string HostApp => GetVariable<string>(HOST_APP);

    public bool IsRunspace => this._ps.Runspace != null;
    public bool IsRunspacePool => this._ps.RunspacePool != null;

    public bool IsDisposed { get; private set; }

    #endregion


    #region Initialization

    public HostedRunspace(PowerShell ps, 
      ILogger? logger = null!, RSLogType rSLogType = RSLogType.Bulk, string logSeparator = "",
      string? hostApp = null!, Encoding encoding = null!)
    {
      this._logger = logger;
      this.Encoding = encoding ?? Encoding.UTF8;
      this.RSLogType = rSLogType;
      this.LogSeparator = logSeparator;

      this._ps = ps;
      Initialize(hostApp);
    }


    public HostedRunspace(
      Dictionary<string, string>? parameters = null!, IEnumerable<string>? modulesToLoad = null!, 
      ILogger? logger = null!, RSLogType rSLogType = RSLogType.Bulk, string logSeparator = "",
      string? hostApp = null!, Encoding encoding = null!, bool restricted = false)
    {
      this._logger = logger;
      this.Encoding = encoding ?? Encoding.UTF8;
      this.RSLogType = rSLogType;
      this.LogSeparator = logSeparator;

      var iss = HostedRunspace.GetInitialSessionState(restricted, parameters, modulesToLoad);

      var rs = RunspaceFactory.CreateRunspace(iss);
      rs.ApartmentState = ApartmentState.MTA;
      rs.ThreadOptions = PSThreadOptions.ReuseThread;
      rs.Open();

      this._ps = PowerShell.Create(rs);
      Initialize(hostApp);
    }


    private void Initialize(string? hostApp)
    {
      SetVariable(HOST_APP, hostApp ?? GetDefaultHostApp(), true, true);

      this.MaxLogLevel = LogLevel.Trace;
#pragma warning disable CS8622 
      this._ps.Streams.Debug.DataAdded += LogMessage<DebugRecord>;
      this._ps.Streams.Verbose.DataAdded += LogMessage<VerboseRecord>;
      this._ps.Streams.Information.DataAdded += LogMessage<InformationRecord>;
      this._ps.Streams.Warning.DataAdded += LogMessage<WarningRecord>;
      this._ps.Streams.Error.DataAdded += LogMessage<ErrorRecord>;
#pragma warning restore CS8622
    }


    internal static InitialSessionState GetInitialSessionState(bool restricted,
      Dictionary<string, string>? parameters = null!, IEnumerable<string>? modulesToLoad = null)
    {
      var iss = InitialSessionState.CreateDefault();
      iss.ExecutionPolicy = restricted ?
        Microsoft.PowerShell.ExecutionPolicy.Restricted :
        Microsoft.PowerShell.ExecutionPolicy.Unrestricted;

      LoadLocalAssembly(iss);

      if (modulesToLoad != null)
      {
        foreach (var moduleName in modulesToLoad)
        {
          iss.ImportPSModule(moduleName);
        }
      }

      if (parameters != null)
      {
        foreach (var parameter in parameters)
        {
          bool constant = parameter.Key.StartsWith("!");
          string name = parameter.Key.StartsWith("!") ?
            parameter.Key[1..] : parameter.Key;

          iss.Variables.Add(new SessionStateVariableEntry(name, parameter.Value, null, GetScopedItemOptions(constant, false)));
        }
      }

      return iss;
    }


    internal static void LoadLocalAssembly(InitialSessionState iss)
    {
      var assemblies = AppDomain.CurrentDomain.GetAssemblies();
      foreach (var assembly in assemblies)
      {
        var types = assembly.GetTypes()
          .Where(c => c.GetCustomAttributes(typeof(CmdletAttribute), true).Length > 0)
          .ToList();
        foreach (var type in types)
        {
          var cmdletattr = type.GetCustomAttributes(typeof(CmdletAttribute), true)
            .Cast<CmdletAttribute>().FirstOrDefault();
          if (cmdletattr != null)
          {
            var ssce = new SessionStateCmdletEntry(
              $"{cmdletattr.VerbName}-{cmdletattr.NounName}", type, null!);
            iss.Commands.Add(ssce);
          }
        }
      }
    }


    public void Dispose()
    {
      this.IsDisposed = true;
      this._ps.Dispose();
      GC.SuppressFinalize(this);
    }


    public static string GetDefaultHostApp() => System.Reflection.Assembly
      .GetEntryAssembly()?.GetName().Name ?? "HostedRuntime";

    #endregion


    #region Events

    public void ClearLog()
    {
      this.Log?.Clear();
      this.MaxLogLevel = LogLevel.Trace;

      this._ps.Streams.Debug.Clear();
      this._ps.Streams.Verbose.Clear();
      this._ps.Streams.Information.Clear();
      this._ps.Streams.Warning.Clear();
      this._ps.Streams.Error.Clear();
    }

    private void LogMessage<T>(object sender, DataAddedEventArgs e)
    {
      var data = ((PSDataCollection<T>)sender)![e.Index];
      var logLevel = LogLevel.Trace;
      if (typeof(T) == typeof(VerboseRecord))
      {
        logLevel = LogLevel.Trace;
      }
      else if (typeof(T) == typeof(DebugRecord))
      {
        logLevel = LogLevel.Debug;
      }
      else if (typeof(T) == typeof(InformationRecord))
      {
        logLevel = LogLevel.Information;
      }
      else if (typeof(T) == typeof(WarningRecord))
      {
        logLevel = LogLevel.Warning;
      }
      else if (typeof(T) == typeof(ErrorRecord))
      {
        logLevel = LogLevel.Error;
      }

      if (this.RSLogType == RSLogType.Direct)
      {
        this._logger?.Log(logLevel, Convert.ToString(data));
      }
      else
      {
        this.Log ??= new StringBuilder();

        this.Log?.Append($@"{this.LogSeparator}""{DateTime.Now}"";""{logLevel}"";" +
          $@"""{Convert.ToString(data)}""");
        if (data != null && data is ErrorRecord er &&
          (er?.InvocationInfo?.PositionMessage?.Length ?? 0) > 0)
        {
          this.Log?.Append($"{Environment.NewLine}{er?.InvocationInfo.PositionMessage}");
        }
        this.Log?.AppendLine();
        this.MaxLogLevel = (LogLevel)(Math.Max((int)logLevel, (int)this.MaxLogLevel));
      }
    }

    private void LogMessage()
    {
      if (this.RSLogType == RSLogType.Bulk)
      {
        this._logger?.Log(this.MaxLogLevel, this.Log?.ToString().Trim());
      }
    }


    #endregion


    #region Functions

    #region Load Module

    public void LoadModule(string moduleName)
    {
      //IsModule(moduleName);
      RunScript($"Import-Module {moduleName} -DisableNameChecking");
    }


    public async Task LoadModuleAsync(string moduleName)
    {
      IsModule(moduleName);
      await RunScriptAsync($"Import-Module {moduleName} -DisableNameChecking");
    }


    private static void IsModule(string moduleName)
    {
      if (moduleName.Contains(' '))
      { throw new ArgumentException($"moduleName cannot have whitespaces: '{moduleName}'"); }
    }

    #endregion


    #region Run


    public async Task<IEnumerable<object?>?> RunAsync(
      CancellationToken cancellationToken = default)
    {
      try
      {
        ClearLog();

        ICollection<PSObject> pipelineObjects;
        if(this.IsRunspacePool)
        {
          pipelineObjects = this._ps.Invoke();
        } else
        {
          //var task = Task.Factory
          //  .FromAsync(this._ps.BeginInvoke(), p => this._ps.EndInvoke(p));
          //await Task.Run(() => task.Wait(), cancellationToken);

          pipelineObjects = await Task.Run<PSDataCollection<PSObject>>(() =>
          {
            var invocation = this._ps.BeginInvoke();
            WaitHandle.WaitAny(new[] {
              invocation.AsyncWaitHandle, cancellationToken.WaitHandle });
          
            if (cancellationToken.IsCancellationRequested)
            {
              this._ps.Stop();
            }
          
            cancellationToken.ThrowIfCancellationRequested();
            return this._ps.EndInvoke(invocation);
          }, cancellationToken);
        }




        return pipelineObjects?.Select(p => p?.BaseObject);
      }
      catch
      { throw; }
      finally
      {
        LogMessage();
      }
    }


    public IEnumerable<object?>? Run()
    {
      try
      {
        ClearLog();
        var pipelineObjects = this._ps.Invoke();
        return pipelineObjects?.Select(p => p?.BaseObject);
      }
      catch
      { throw; }
      finally
      {
        LogMessage();
      }
    }

    #endregion


    #region Command

    /// <summary>
    /// Adds a command to the Powershell runtime
    /// </summary>
    /// <param name="command">Command to be added</param>
    public void AddCommand(string command, Dictionary<string, object?>? parameters = null!,
      bool clearPipeline = false)
    {
      if (clearPipeline)
      {
        this._ps.Commands.Clear();
      }

      this._ps.AddCommand(command);
      if (parameters != null && parameters.Count > 0)
      {
        var pf = GetParametersFromFunction(command);
        if (pf == null)
        { throw new CommandNotFoundException(command); }

        var parametersFiltered = parameters
          .Where(p => p.Value != null!)
          .Where(p => pf.ContainsKey(p.Key) &&
            (p.Value != null && (p.Value.GetType().IsAssignableTo(pf[p.Key]) ||
             pf[p.Key].IsAssignableFrom(typeof(System.Collections.Hashtable)) &&
             typeof(System.Collections.IDictionary).IsAssignableFrom(p.Value.GetType()))))
          .ToDictionary(k => k.Key, v => v.Value, StringComparer.OrdinalIgnoreCase);
        this._ps.AddParameters(parametersFiltered);
      }
    }


    /// <summary>
    /// Runs a PowerShell script with parameters and prints the resulting pipeline 
    /// objects to the console output. 
    /// </summary>
    /// <param name="scriptContents">The script file contents.</param>
    /// <param name="scriptParameters">A dictionary of parameter names and parameter 
    /// values.</param>
    public async Task<IEnumerable<object?>?> RunCommandAsync(string command,
      Dictionary<string, object?>? parameters = null!, bool clearPipeline = true,
      CancellationToken cancellationToken = default)
    {
      AddCommand(command, parameters, clearPipeline);
      return await RunAsync(cancellationToken);
    }


    /// <summary>
    /// Runs a PowerShell script with parameters and prints the resulting pipeline 
    /// objects to the console output. 
    /// </summary>
    /// <param name="scriptContents">The script file contents.</param>
    /// <param name="scriptParameters">A dictionary of parameter names and parameter 
    /// values.</param>
    public IEnumerable<object?>? RunCommand(string command,
      Dictionary<string, object?>? parameters = null!, bool clearPipeline = true)
    {
      AddCommand(command, parameters, clearPipeline);
      return Run();
    }

    #endregion


    #region Script

    /// <summary>
    /// Adds a command to the Powershell runtime
    /// </summary>
    /// <param name="command">Command to be added</param>
    public void AddScript(string script, bool clearPipeline = false)
    {
      if (clearPipeline)
      { this._ps.Commands.Clear(); }

      this._ps.AddScript(script);
    }


    /// <summary>
    /// Runs a PowerShell script with parameters and prints the resulting pipeline 
    /// objects to the console output. 
    /// </summary>
    /// <param name="scriptContents">The script file contents.</param>
    /// <param name="scriptParameters">A dictionary of parameter names and parameter 
    /// values.</param>
    public async Task<IEnumerable<object?>?> RunScriptAsync(string script,
      bool clearPipeline = true, CancellationToken cancellationToken = default)
    {
      AddScript(script, clearPipeline);
      return await RunAsync(cancellationToken);
    }


    /// <summary>
    /// Runs a PowerShell script with parameters and prints the resulting pipeline 
    /// objects to the console output. 
    /// </summary>
    /// <param name="scriptContents">The script file contents.</param>
    /// <param name="scriptParameters">A dictionary of parameter names and parameter 
    /// values.</param>
    public IEnumerable<object?>? RunScript(string script,
      bool clearPipeline = true)
    {
      AddScript(script, clearPipeline);
      return Run();
    }


    /// <summary>
    /// Loads a script to a runtime
    /// </summary>
    /// <param name="script">Script to be executed</param>
    public async Task<IEnumerable<object?>?> RunScriptFromFileAsync(string filepath,
      Encoding? encoding = null!, CancellationToken cancellationToken = default)
    {
      return await RunScriptAsync(
        System.IO.File.ReadAllText(filepath, encoding ?? this.Encoding), false,
        cancellationToken);
    }


    /// <summary>
    /// Loads a script to a runtime
    /// </summary>
    /// <param name="script">Script to be executed</param>
    public IEnumerable<object?>? RunScriptFromFile(string filepath,
      Encoding? encoding = null!)
    {
      return RunScript(
        System.IO.File.ReadAllText(filepath, encoding ?? this.Encoding), false);
    }


    /// <summary>
    /// Loads scripts to a runtime
    /// </summary>
    /// <param name="script">Script to be executed</param>
    public IEnumerable<object?>? RunScriptFromFile(IEnumerable<string> filepaths,
      Encoding? encoding = null!)
    {
      foreach (var filepath in filepaths)
      {
        AddScript(System.IO.File.ReadAllText(filepath, encoding ?? this.Encoding), false);
      }
      return Run();
    }


    /// <summary>
    /// Loads scripts to a runtime
    /// </summary>
    /// <param name="script">Script to be executed</param>
    public async Task<IEnumerable<object?>?> RunScriptFromFileAsync(IEnumerable<string> filepaths,
      Encoding? encoding = null!, CancellationToken cancellationToken = default)
    {
      foreach (var filepath in filepaths)
      {
        AddScript(System.IO.File.ReadAllText(filepath, encoding ?? this.Encoding), false);
      }
      return await RunAsync(cancellationToken);
    }


    /// <summary>
    /// Loads a script to a runtime
    /// </summary>
    /// <param name="script">Script to be executed</param>
    public IEnumerable<object?>? RunScriptFromPath(string path,
      Encoding? encoding = null!, string searchPattern = "*.ps1",
      System.IO.SearchOption searchOption = System.IO.SearchOption.TopDirectoryOnly)
    {
      if (path == null)
      { return null!; }

      if (System.IO.File.Exists(path))
      {
        return RunScriptFromFile(path, encoding);
      }

      foreach (var filepath in System.IO.Directory.GetFiles(path, searchPattern,
        searchOption))
      {
        AddScript(System.IO.File.ReadAllText(filepath, encoding ?? this.Encoding), false);
      }
      return Run();
    }


    /// <summary>
    /// Loads a script to a runtime
    /// </summary>
    /// <param name="script">Script to be executed</param>
    public async Task<IEnumerable<object?>?> RunScriptFromPathAsync(string path,
      Encoding? encoding = null!, string searchPattern = "*.ps1",
      System.IO.SearchOption searchOption = System.IO.SearchOption.TopDirectoryOnly,
      CancellationToken cancellationToken = default)
    {
      if (path == null)
      { return null!; }

      if (System.IO.File.Exists(path))
      {
        this._logger?.LogInformation($"Running script from file: '{path}'");
        return await RunScriptFromFileAsync(path, encoding, cancellationToken: cancellationToken);
      }

      var log = new StringBuilder();
      foreach (var filepath in System.IO.Directory.GetFiles(path, searchPattern,
        searchOption))
      {
        log.AppendLine($"-> Loading file: '{filepath}'");
        AddScript(System.IO.File.ReadAllText(filepath, encoding ?? this.Encoding), false);
      }
      this._logger?.LogInformation($"Running scripts from file: {Environment.NewLine}{log}");
      return await RunAsync(cancellationToken);
    }

    #endregion


    #region Variables/Runspace

    public Dictionary<string, Type>? GetParametersFromFunction(string command)
    {
      IEnumerable<CommandInfo>? commands;
      if (this.IsRunspace)
      {
        commands = this._ps.Runspace.SessionStateProxy.InvokeCommand
          .GetCommands(command, CommandTypes.All, false);
      }
      else
      {
        commands = RunScript($"Get-Command -Name {command}", true)?
          .Cast<CommandInfo>();
      }

      return commands?.FirstOrDefault()?
        .Parameters
        .ToDictionary(
          k => k.Key,
          v => v.Value.ParameterType,
          StringComparer.OrdinalIgnoreCase
          );
    }


    public bool ContainsParameter(string command, string parameter)
    {
      return GetParametersFromFunction(command)?
        .ContainsKey(parameter) ?? false;
    }


    public Dictionary<string, Dictionary<string, Type>> GetFunctions()
    {
      IEnumerable<CommandInfo> commandInfo = null!;
      if (this.IsRunspace)
      {
        commandInfo = this._ps.Runspace.SessionStateProxy.InvokeCommand
          .GetCommands("*", CommandTypes.Function, true);
      }
      else
      {
        var pipe = RunScript($"Get-Command", true)?
          .Cast<CommandInfo>();
      }

      return commandInfo
        .ToDictionary(
          k => k.Name,
          v => v.Parameters.ToDictionary(
            kp => kp.Key,
            vp => vp.Value.ParameterType,
            StringComparer.OrdinalIgnoreCase));
    }


    private static ScopedItemOptions GetScopedItemOptions(bool constant = false, bool allScope = false)
    {
      ScopedItemOptions scopedItemOptions = ScopedItemOptions.None;
      if (constant)
      {
        scopedItemOptions |= ScopedItemOptions.Constant;
      }

      if (allScope)
      {
        scopedItemOptions |= ScopedItemOptions.AllScope;
      }
      return scopedItemOptions;
    }


    private static PSVariable GetPSVariable(string name, object value, bool constant, bool global)
    {
      if (global && !name.StartsWith("Global:", StringComparison.OrdinalIgnoreCase))
      {
        name = $"Global:{name}";
      }

      return new PSVariable(name, value, GetScopedItemOptions(constant, false));
    }


    public void SetVariable(string name, object value, bool constant = false,
      bool global = false)
    {
      if (this.IsRunspace)
      {
        this._ps.Runspace.SessionStateProxy.PSVariable.Set(
          GetPSVariable(name, value, constant, global));
      } else
      {
        RunCommand("Set-Variable", new Dictionary<string, object?>() 
        {
          { "Name", name },
          { "Value", value },
          { "Option", GetScopedItemOptions(constant, global) }
        });
      }
    }


    public object GetVariable(string name)
    {
      if (this.IsRunspace)
      {
        return this._ps.Runspace.SessionStateProxy.GetVariable(name);
      }
      else
      {
        return RunScript("${name}")?.FirstOrDefault()!;
      }
    }


    public T GetVariable<T>(string name)
    {
      return (T)GetVariable(name);
    }

    #endregion

    #endregion

  }
}
