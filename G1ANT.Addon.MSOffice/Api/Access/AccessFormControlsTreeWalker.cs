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
    internal class AccessFormControlsTreeWalker : IAccessFormControlsTreeWalker
    {
        public AccessControlModel GetAccessControlByPath(Application application, string path)
        {
            return GetAccessControlByPath(application, new ControlPathModel(path));
        }

        public AccessControlModel GetAccessControlByPath(Application application, ControlPathModel controlPath)
        {
            var form = application.Forms[controlPath.FormName];

            _Control controlFound = null;

            var controls = form.Controls.OfType<_Control>().ToList();

            for (var i = 0; i < controlPath.Count; i++)
            {
                var pathElement = controlPath[i];
                controlFound = GetMatchingControl(controls, controlPath[i]);

                if (i == controlPath.Count - 1)
                    break; // don't load children of last element of path, it's probably not a container and wil throw a COM exception
                controls = controlFound.Controls.OfType<_Control>().ToList();
            }

            return new AccessControlModel(controlFound, true, false);
        }


        private _Control GetMatchingControl(ICollection<_Control> controls, ControlPathElementModel pathElement)
        {

            IEnumerable<_Control> controlsFound = controls;
            if (pathElement.ShouldFilterByPropertyNameAndValue())
                controlsFound = controlsFound.Where(c => c.Properties[pathElement.PropertyName]?.Value?.ToString() == pathElement.PropertyValue);

            if (pathElement.ShouldFilterByIndex())
                controlsFound = controlsFound.Skip(pathElement.ChildIndex);

            var controlFound = controlsFound.FirstOrDefault();

            if (controlFound == null)
                throw new Exception($"Element {pathElement} not found (name={pathElement.PropertyName}, value={pathElement.PropertyValue}, index={pathElement.ChildIndex})");

            return controlFound;
        }

    }
}
