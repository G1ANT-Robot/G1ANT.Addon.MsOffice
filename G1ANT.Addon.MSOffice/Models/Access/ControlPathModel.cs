using System;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class ControlPathModel : List<string>
    {
        public const char PathSeparator = '/';
        public string Path { get; }
        public string FormName { get; }

        public ControlPathModel(string path)
        {
            Path = path;
            var pathElements = path.Split(new[] { PathSeparator }, StringSplitOptions.RemoveEmptyEntries);

            FormName = pathElements[0];
            AddRange(pathElements.Skip(1));
        }
    }
}
