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
            var controlPath = new ControlPathModel(path);

            var form = application.Forms[controlPath.FormName];

            Control controlFound = null;

            var controls = form.Controls.OfType<Control>().ToList();

            for (var i = 0; i < controlPath.Count; i++)
            {
                var pathElement = controlPath[i];
                controlFound = GetMatchingControl(controls, controlPath[i]);

                if (i == controlPath.Count - 1)
                    break; // don't load children of last element of path, it's probably not a container and wil throw a COM exception
                controls = controlFound.Controls.OfType<Control>().ToList();
            }

            return new AccessControlModel(controlFound, true, false);
        }


        private Control GetMatchingControl(ICollection<Control> controls, string pathElement)
        {
            var element = new ControlPathElementModel(pathElement);

            IEnumerable<Control> controlsFound = controls;
            if (element.ShouldFilterByPropertyNameAndValue())
                controlsFound = controlsFound.Where(c => c.Properties[element.PropertyName]?.Value?.ToString() == element.PropertyValue);

            if (element.ShouldFilterByIndex())
                controlsFound = controlsFound.Skip(element.ChildIndex);

            var controlFound = controlsFound.FirstOrDefault();

            if (controlFound == null)
                throw new Exception($"Element {element} not found (name={element.PropertyName}, value={element.PropertyValue}, index={element.ChildIndex})");

            return controlFound;
        }

    }
}
