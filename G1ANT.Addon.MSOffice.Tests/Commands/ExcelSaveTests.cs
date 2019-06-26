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
    [Apartment(ApartmentState.STA)]
    public class ExcelSaveTests
	{
        Scripter scripter;
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

        private int ExcelProcessesCount()
        {
            return Process.GetProcessesByName("excel").Length;
        }
        [OneTimeSetUp]
        public void ClassInit()
        {
            Environment.CurrentDirectory = TestContext.CurrentContext.TestDirectory;
            scripter = new Scripter();
scripter.InitVariables.Clear();
        }
        [SetUp]
        public void init()
        {
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.MSOffice.dll");
        }
        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
		public void ExcelSaveTest()
		{
            try
            {
                string saveDir = System.IO.Directory.GetCurrentDirectory() + @"\";
                string savePath = saveDir + "test.xlsx";
                FileInfo savedFile = new FileInfo(savePath);
                if (savedFile.Exists)
                {
                    savedFile.Delete();
                }
               scripter.InitVariables.Add("savePath", new TextStructure(savePath));

                scripter.RunLine("excel.open");
                scripter.RunLine("excel.addsheet test1");
                scripter.RunLine("excel.activatesheet test1");
                scripter.RunLine("excel.setvalue 3 row 1 colindex 1");
                scripter.RunLine("excel.addsheet test2");
                scripter.RunLine("excel.activatesheet test2");
                scripter.RunLine("excel.setvalue 5 row 2 colindex 1");
                scripter.RunLine($"excel.save {SpecialChars.Variable}savePath");
                scripter.RunLine("excel.close");

                scripter.RunLine($"excel.open {SpecialChars.Variable}savePath");
                scripter.RunLine("excel.activatesheet test1");
                scripter.RunLine("excel.getvalue row 1 colindex 1");
                Assert.AreEqual(3, int.Parse(scripter.Variables.GetVariableValue<string>("result")));
                scripter.RunLine("excel.activatesheet test2");
                scripter.RunLine("excel.getvalue row 2 colindex 1");
                Assert.AreEqual(5, int.Parse(scripter.Variables.GetVariableValue<string>("result")));
                scripter.RunLine("excel.close");
            }
            catch (Exception ex)
            {
                KillProcesses();
                throw ex.GetBaseException();
            }
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
