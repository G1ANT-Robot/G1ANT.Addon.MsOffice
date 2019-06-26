/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Addon.MSOffice;

using System;
using System.IO;
using NUnit.Framework;
using System.Threading;

using System.Reflection;
using System.Diagnostics;
using G1ANT.Engine;
using G1ANT.Addon.MSOffice.Tests.Properties;
using G1ANT.Language;

namespace G1ANT.Addon.MSOffice.Tests
{
    [TestFixture]
    [Apartment(ApartmentState.STA)]
    public class ExcelDuplicateRowTests
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
        public void ExcelDuplicateRowTest()
        {
            scripter.Text = ($@"excel.open ♥xlsPath sheet Add
                                excel.duplicaterow source 1 destination 2
                                excel.getvalue row 2 colindex 1
                                excel.close");
            scripter.Run();
            Assert.AreEqual(3, int.Parse(scripter.Variables.GetVariableValue<string>("result")));
            
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelDuplicateRowFailTest()
        {
            scripter.Text = ($@"excel.open ♥xlsPath sheet Add
                              excel.duplicaterow source 1 destination 2
                              excel.getvalue row 0 colindex 1");

            Exception exception = Assert.Throws<ApplicationException>(delegate
            {
                scripter.Run();
            });
            Assert.IsInstanceOf<ArgumentException>(exception.GetBaseException());

        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelDuplicateRowFail2Test()
        {
            scripter.Text = ($@"excel.open ♥xlsPath sheet Add
                                excel.duplicaterow source 1 destination 2
                                excel.getvalue row -1 colindex -1");
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
