// FILE: src/Integrations/HardwareProbeAdapter.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Read hardware identifiers via System.Management for licensing.
//   SCOPE: Produces stable machine identifier based on WMI data.
//   DEPENDS: none
//   LINKS: M-HARDWARE-PROBE
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   GetHardwareId - Retrieves machine hardware fingerprint.
// END_MODULE_MAP

// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.0.0 - Added missing CHANGE_SUMMARY block for GRACE integrity refresh.
// END_CHANGE_SUMMARY


using System.Management;

namespace ElTools.Integrations;

public class HardwareProbeAdapter
{
    // START_CONTRACT: GetHardwareId
    //   PURPOSE: Retrieve hardware id.
    //   INPUTS: none
    //   OUTPUTS: { string - textual result for retrieve hardware id }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-HARDWARE-PROBE
    // END_CONTRACT: GetHardwareId

    public string GetHardwareId()
    {
        // START_BLOCK_GET_HARDWARE_ID
        try
        {
            using var searcher = new ManagementObjectSearcher("SELECT ProcessorId FROM Win32_Processor");
            foreach (ManagementObject obj in searcher.Get())
            {
                string? value = obj["ProcessorId"]?.ToString();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }
            }
        }
        catch
        {
        }

        return "UNKNOWN-HWID";
        // END_BLOCK_GET_HARDWARE_ID
    }
}