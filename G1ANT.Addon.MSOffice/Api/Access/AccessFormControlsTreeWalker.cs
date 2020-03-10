/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using G1ANT.Addon.MSOffice.Models.Access;
using Microsoft.Office.Interop.Access;
using System;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Access
{

    public class AccessFormControlsTreeWalker : IAccessFormControlsTreeWalker
    {
        public AccessControlModel GetAccessControlByPath(Application application, string path)
        {
            var pathElements = path.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);

            var formName = pathElements[0];
            var form = application.Forms[formName];

            Control controlFound = null;

            var controls = form.Controls.OfType<Control>().ToList();

            pathElements = pathElements.Skip(1).ToArray();

            for (var i = 0; i < pathElements.Length; i++)
            {
                controlFound = GetMatchingControl(controls, pathElements[i]);

                if (i == pathElements.Length - 1)
                    break; // don't load children of last element of path, it's probably not a container and wil throw a COM exception
                controls = controlFound.Controls.OfType<Control>().ToList();
            }

            return new AccessControlModel(controlFound, true, false);
        }


        private Control GetMatchingControl(ICollection<Control> controls, string pathElement)
        {
            if (string.IsNullOrEmpty(pathElement))
                throw new ArgumentNullException(nameof(pathElement));

            var pathParts = pathElement.Split('=');
            var propertyName = pathParts.Length > 1 ? pathParts[0] : "Name";
            var propertyValue = pathParts.Last();
            var childIndex = -1;

            if (propertyValue.Contains("[") && propertyValue.Contains("]"))
            {
                var childIndexValue = propertyValue.Substring(propertyValue.IndexOf("[") + 1);
                childIndexValue = childIndexValue.Substring(0, childIndexValue.IndexOf("]"));
                childIndex = int.Parse(childIndexValue);
                propertyValue = propertyValue.Substring(0, propertyValue.IndexOf("["));
            }

            IEnumerable<Control> controlsFound = controls;
            if (pathParts.Contains("=") || childIndex < 0)
                controlsFound = controlsFound.Where(c => c.Properties[propertyName]?.Value?.ToString() == propertyValue);

            if (childIndex >= 0)
                controlsFound = controlsFound.Skip(childIndex);

            var controlFound = controlsFound.FirstOrDefault();

            if (controlFound == null)
                throw new Exception($"Element {pathElement} not found (index is {childIndex})");

            return controlFound;
        }

    }
}
