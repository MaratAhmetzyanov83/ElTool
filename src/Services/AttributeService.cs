// FILE: src/Services/AttributeService.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Read and write block attributes with robust Cyrillic support.
//   SCOPE: Attribute extraction and update for BlockReference entities.
//   DEPENDS: M-CAD-CONTEXT
//   LINKS: M-ATTRIBUTES, M-CAD-CONTEXT
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   ReadAttributes - Returns attribute dictionary from block.
//   WriteAttributes - Writes attribute dictionary to block.
// END_MODULE_MAP

// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.0.0 - Added missing CHANGE_SUMMARY block for GRACE integrity refresh.
// END_CHANGE_SUMMARY


using ElTools.Integrations;

namespace ElTools.Services;

public class AttributeService
{
    private readonly AutoCADAdapter _acad = new();

    public Dictionary<string, string> ReadAttributes(ObjectId blockId)
    {
        // START_BLOCK_READ_ATTRIBUTES
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        _acad.RunTransaction((tr, _) =>
        {
            var block = tr.GetObject(blockId, OpenMode.ForRead) as BlockReference;
            if (block is null)
            {
                return;
            }

            foreach (ObjectId attId in block.AttributeCollection)
            {
                if (tr.GetObject(attId, OpenMode.ForRead) is AttributeReference att)
                {
                    result[att.Tag] = att.TextString;
                }
            }
        });

        return result;
        // END_BLOCK_READ_ATTRIBUTES
    }

    // START_CONTRACT: WriteAttributes
    //   PURPOSE: Write attributes.
    //   INPUTS: { blockId: ObjectId - method parameter; attributes: IReadOnlyDictionary<string, string> - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-ATTRIBUTES
    // END_CONTRACT: WriteAttributes

    public void WriteAttributes(ObjectId blockId, IReadOnlyDictionary<string, string> attributes)
    {
        // START_BLOCK_WRITE_ATTRIBUTES
        _acad.RunTransaction((tr, _) =>
        {
            var block = tr.GetObject(blockId, OpenMode.ForWrite) as BlockReference;
            if (block is null)
            {
                return;
            }

            foreach (ObjectId attId in block.AttributeCollection)
            {
                if (tr.GetObject(attId, OpenMode.ForWrite) is not AttributeReference att)
                {
                    continue;
                }

                if (attributes.TryGetValue(att.Tag, out string? value))
                {
                    att.TextString = value;
                }
            }
        });
        // END_BLOCK_WRITE_ATTRIBUTES
    }
}