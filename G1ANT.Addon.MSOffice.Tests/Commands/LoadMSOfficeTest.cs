/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using G1ANT.Engine;
using G1ANT.Language;
using NUnit.Framework;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace G1ANT.Addon.MSOffice.Tests
{
    [TestFixture]
    class LoadMSOfficeTest
    {
        [Test]
        public void LoadMSOfficeAddon()
        {
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.Mscrm.dll");
            Assert.IsNotNull(addon);
            Assert.IsTrue(addon.Commands.Count > 0);
        }
    }
}
