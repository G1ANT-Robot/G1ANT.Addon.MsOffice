/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using G1ANT.Language;
using Microsoft.Office.Interop.Outlook;
using System;

namespace G1ANT.Addon.MSOffice
{
    [Structure(Name = "OutlookAttachment", AutoCreate = false)]
    public class OutlookAttachmentStructure : StructureTyped<Attachment>
    {
        const string FilenameIndex = "filename";
        const string SizeIndex = "size";

        public OutlookAttachmentStructure(string value, string format = "", AbstractScripter scripter = null) :
            base(value, format, scripter)
        {
            Init();
        }

        public OutlookAttachmentStructure(object value, string format = null, AbstractScripter scripter = null)
            : base(value, format, scripter)
        {
            Init();
        }

        protected void Init()
        {
            Indexes.Add(FilenameIndex);
            Indexes.Add(SizeIndex);
        }

        public override Structure Get(string index = "")
        {
            if (string.IsNullOrWhiteSpace(index))
                return new OutlookAttachmentStructure(Value, Format);
            switch (index.ToLower())
            {
                case FilenameIndex:
                    return new TextStructure(Value.FileName, null, Scripter);
                case SizeIndex:
                    return new IntegerStructure(Value.Size, null, Scripter);
            }
            throw new ArgumentException($"Unknown index '{index}'");
        }

        public override void Set(Structure structure, string index = null)
        {
            if (structure == null || structure.Object == null)
                throw new ArgumentNullException(nameof(structure));
            else
                throw new ArgumentException($"Unknown index '{index}'");
        }

        public override string ToString(string format)
        {
            return Value?.ToString();
        }

        protected override Attachment Parse(string value, string format = null)
        {
            return null;
        }
    }
}
