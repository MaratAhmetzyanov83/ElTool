// FILE: src/Services/InstallTypeResolver.cs
// VERSION: 1.1.0
// START_MODULE_CONTRACT
//   PURPOSE: Resolve install type by linetype/layer rules with deterministic priority and default fallback.
//   SCOPE: Runtime matching for EOM_TRACE aggregation and validation.
//   DEPENDS: M-SETTINGS, M-CONFIG, M-MODELS
//   LINKS: M-INSTALL-TYPE-RESOLVER, M-SETTINGS, M-CONFIG, M-MODELS, M-CABLE-CALC
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   Resolve - Resolves first matching rule ordered by Priority.
// END_MODULE_MAP

using ElTools.Models;
using ElTools.Shared;

namespace ElTools.Services;

public interface IInstallTypeResolver
{
    string Resolve(string linetypeResolved, string layerName, InstallTypeRuleSet rules);
}

public sealed class InstallTypeResolver : IInstallTypeResolver
{
    public string Resolve(string linetypeResolved, string layerName, InstallTypeRuleSet rules)
    {
        // START_BLOCK_RESOLVE_INSTALL_TYPE
        foreach (InstallTypeRule rule in rules.Rules.OrderBy(x => x.Priority))
        {
            bool matches = rule.MatchBy.Equals("Linetype", StringComparison.OrdinalIgnoreCase)
                ? string.Equals(linetypeResolved, rule.Value, StringComparison.OrdinalIgnoreCase)
                : rule.MatchBy.Equals("Layer", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(layerName, rule.Value, StringComparison.OrdinalIgnoreCase);

            if (matches)
            {
                return rule.Result;
            }
        }

        return string.IsNullOrWhiteSpace(rules.Default) ? PluginConfig.Strings.Unknown : rules.Default;
        // END_BLOCK_RESOLVE_INSTALL_TYPE
    }
}

