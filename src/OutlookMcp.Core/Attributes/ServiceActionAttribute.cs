namespace OutlookMcp.Core.Attributes;

/// <summary>
/// Overrides the default action name derived from method name.
/// By default, action names are derived from method names using PascalCase → kebab-case convention.
/// Use this attribute only when the convention doesn't produce the desired action name.
/// </summary>
/// <remarks>
/// Convention: GetLoadConfig → "get-load-config"
/// Override example: [ServiceAction("custom-action")]
/// </remarks>
[AttributeUsage(AttributeTargets.Method, AllowMultiple = false, Inherited = false)]
public sealed class ServiceActionAttribute : Attribute
{
    /// <summary>
    /// The action name to use instead of the derived name.
    /// </summary>
    public string Action { get; }

    /// <summary>
    /// Overrides the tool-level [McpTool] Destructive default for this specific action.
    /// Use this when a tool exposes a mix of read-only and mutating actions (action-dispatch
    /// tools cannot be described by a single tool-level boolean). The generator computes the
    /// tool's overall Destructive hint as true if ANY action is destructive.
    /// Attribute parameters cannot be nullable, so absence of this named argument (not whether
    /// it equals a particular value) is what signals "no override" to the generator.
    /// </summary>
    public bool Destructive { get; set; }

    /// <summary>
    /// Creates a new ServiceActionAttribute.
    /// </summary>
    /// <param name="action">The action name in kebab-case (e.g., "get-load-config")</param>
    public ServiceActionAttribute(string action)
    {
        Action = action ?? throw new ArgumentNullException(nameof(action));
    }
}
