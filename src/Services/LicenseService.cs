// FILE: src/Services/LicenseService.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Validate license state before command execution.
//   SCOPE: Reads local settings and validates HWID policy.
//   DEPENDS: M-HARDWARE-PROBE, M-SETTINGS, M-LOGGING, M-MODELS
//   LINKS: M-LICENSE, M-HARDWARE-PROBE, M-SETTINGS, M-LOGGING, M-MODELS
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   Validate - Returns current license state.
// END_MODULE_MAP

// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.0.0 - Added missing CHANGE_SUMMARY block for GRACE integrity refresh.
// END_CHANGE_SUMMARY


using ElTools.Data;
using ElTools.Integrations;
using ElTools.Models;

namespace ElTools.Services;

public class LicenseService
{
    private readonly HardwareProbeAdapter _hardware = new();
    private readonly SettingsRepository _settingsRepository = new();
    private readonly LogService _log = new();

    // START_CONTRACT: Validate
    //   PURPOSE: Validate.
    //   INPUTS: none
    //   OUTPUTS: { LicenseState - result of validate }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-LICENSE
    // END_CONTRACT: Validate

    public LicenseState Validate()
    {
        // START_BLOCK_VALIDATE_LICENSE
        string hwid = _hardware.GetHardwareId();
        _ = _settingsRepository.LoadSettings();

        if (hwid == "UNKNOWN-HWID")
        {
            _log.Write("Р вЂєР С‘РЎвЂ Р ВµР Р…Р В·Р С‘РЎР‚Р С•Р Р†Р В°Р Р…Р С‘Р Вµ Р С•РЎвЂљР С”Р В»РЎР‹РЎвЂЎР ВµР Р…Р С•: HWID Р Р…Р ВµР Т‘Р С•РЎРѓРЎвЂљРЎС“Р С—Р ВµР Р….");
            return LicenseState.Disabled;
        }

        _log.Write("Р СџРЎР‚Р С•Р Р†Р ВµРЎР‚Р С”Р В° Р В»Р С‘РЎвЂ Р ВµР Р…Р В·Р С‘Р С‘ Р Р†РЎвЂ№Р С—Р С•Р В»Р Р…Р ВµР Р…Р В°.");
        return LicenseState.Valid;
        // END_BLOCK_VALIDATE_LICENSE
    }
}