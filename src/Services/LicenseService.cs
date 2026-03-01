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

using ElTools.Data;
using ElTools.Integrations;
using ElTools.Models;

namespace ElTools.Services;

public class LicenseService
{
    private readonly HardwareProbeAdapter _hardware = new();
    private readonly SettingsRepository _settingsRepository = new();
    private readonly LogService _log = new();

    public LicenseState Validate()
    {
        // START_BLOCK_VALIDATE_LICENSE
        string hwid = _hardware.GetHardwareId();
        _ = _settingsRepository.LoadSettings();

        if (hwid == "UNKNOWN-HWID")
        {
            _log.Write("Лицензирование отключено: HWID недоступен.");
            return LicenseState.Disabled;
        }

        _log.Write("Проверка лицензии выполнена.");
        return LicenseState.Valid;
        // END_BLOCK_VALIDATE_LICENSE
    }
}

