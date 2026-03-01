// FILE: src/Integrations/JsonAdapter.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Serialize and deserialize plugin settings using Newtonsoft.Json.
//   SCOPE: Safe JSON serialization and deserialization helpers.
//   DEPENDS: none
//   LINKS: M-JSON-ADAPTER
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   Serialize - Converts model to JSON string.
//   Deserialize - Converts JSON string to model.
// END_MODULE_MAP

// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.0.0 - Added missing CHANGE_SUMMARY block for GRACE integrity refresh.
// END_CHANGE_SUMMARY


using Newtonsoft.Json;

namespace ElTools.Integrations;

public class JsonAdapter
{
    public string Serialize<T>(T model)
    {
        // START_BLOCK_SERIALIZE
        return JsonConvert.SerializeObject(model, Formatting.Indented);
        // END_BLOCK_SERIALIZE
    }

    public T? Deserialize<T>(string json)
    {
        // START_BLOCK_DESERIALIZE
        if (string.IsNullOrWhiteSpace(json))
        {
            return default;
        }

        return JsonConvert.DeserializeObject<T>(json);
        // END_BLOCK_DESERIALIZE
    }
}