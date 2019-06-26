/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using System;
using System.IO;
using NUnit.Framework;
using System.Threading;

using System.Reflection;
using System.Diagnostics;
using G1ANT.Engine;
using G1ANT.Language;
using G1ANT.Addon.MSOffice.Tests.Properties;

namespace G1ANT.Addon.MSOffice.Tests
{
    [TestFixture]
    [Apartment(ApartmentState.STA)]
    public class ExcelActivateSheetTests
    {

        Scripter scripter;
        static string xlsPath;

        private void KillProcesses()
        {
            foreach (Process p in Process.GetProcessesByName("excel"))
            {
                try
                {
                    p.Kill();
                }
                catch { }
            }
        }

        [OneTimeSetUp]
        public void ClassInit()
        {
            Environment.CurrentDirectory = TestContext.CurrentContext.TestDirectory;
            
        }

        [SetUp]
        public void TestInit()
        {
            scripter = new Scripter();
            scripter.InitVariables.Clear();
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.MSOffice.dll");
            xlsPath = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.TestWorkbook), "xlsm");  
           scripter.InitVariables.Add("xlsPath", new TextStructure(xlsPath));
           
        }

        

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelActivateSheetTest()
        {
            scripter.Text =($@"excel.open {SpecialChars.Variable}xlsPath
                              excel.activatesheet name {SpecialChars.Text}Macro{SpecialChars.Text}
                              excel.close");
            scripter.Run();
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelActivateSheetFailTest()
        {
            scripter.Text = ($@"excel.open {SpecialChars.Variable}xlsPath
                                excel.activatesheet name {SpecialChars.Text}aaaa{SpecialChars.Text}");
            Exception exception = Assert.Throws<ApplicationException>(delegate
            {
                scripter.Run();
            });
            Assert.IsInstanceOf<ArgumentException>(exception.GetBaseException());
        }

        [TearDown]
        public void TestCleanUp()
        {
            Process[] proc = Process.GetProcessesByName("excel");
            if (proc.Length != 0)
            {
                KillProcesses();
            }
        }
    }
}
