using G1ANT.Language;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;

namespace G1ANT.Addon.MSOffice.Api
{
    internal class StructureConverter
    {
        internal Structure Convert(object value)
        {
            switch (value)
            {
                case Structure structure:
                    return structure;

                case string text:
                    return new TextStructure(text);
                case int integer:
                    return new IntegerStructure(integer);
                case float @float:
                    return new FloatStructure(@float);
                case bool @bool:
                    return new BooleanStructure(@bool);
                case DateTime dateTime:
                    return new DateTimeStructure(dateTime);

                case IEnumerable ienumerable:
                    return new JsonStructure(JArray.FromObject(ienumerable));
                default:
                    return new JsonStructure(JObject.FromObject(value));
            }
        }
    }
}
