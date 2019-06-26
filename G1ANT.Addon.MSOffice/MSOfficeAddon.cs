/**
*    Copyright(C) G1ANT Robot Ltd, All rights reserved
*    Solution G1ANT.Addon.MSOffice, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;
using System.Drawing;

namespace G1ANT.Addon.MSOffice
{
    [Addon(Name = "MSOffice", Tooltip = "MSOffice Commands")]
    [Copyright(Author = "G1ANT Robot Ltd", Copyright = "G1ANT Robot Ltd", Email = "hi@g1ant.com", Website = "www.g1ant.com")]
    [License(Type = "LGPL", ResourceName = "License.txt")]
    [CommandGroup(Name = "excel", Tooltip = "Command connected with creating, editing and generally working on excel")]
    [CommandGroup(Name = "word",  Tooltip = "Command connected with creating, editing and generally working on word")]
    [CommandGroup(Name = "outlook", Tooltip = "Command connected with creating, editing and generally working on outlook")]
    public class MSOfficeAddon : Language.Addon
    {
    }
}