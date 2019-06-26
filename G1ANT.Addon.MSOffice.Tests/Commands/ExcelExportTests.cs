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

using G1ANT.Engine;
using System.Reflection;
using System.Diagnostics;
using G1ANT.Language;
using G1ANT.Addon.MSOffice.Tests.Properties;

namespace G1ANT.Addon.MSOffice.Tests
{
    [TestFixture]
    [Apartment(ApartmentState.STA)]
    public class ExcelExportTests
    {
        static FileInfo excelFile;
        Scripter scripter;
        static string xlsPath, pdfPath, xpsPath, excelPath;

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
            excelPath = Assembly.GetExecutingAssembly().UnpackResourceToFile("Resources." + nameof(Resources.TestWorkbook), "xlsm");

            excelFile = new FileInfo(Path.Combine(Environment.CurrentDirectory, excelPath));
            xlsPath = excelFile.FullName;
            pdfPath = excelFile.DirectoryName + @"\" + excelFile.Name + ".pdf";
            xpsPath = excelFile.DirectoryName + @"\" + excelFile.Name + ".xps";

            scripter = new Scripter();
            scripter.InitVariables.Clear();
            scripter.InitVariables.Add("xlsPath", new TextStructure(xlsPath));
            scripter.InitVariables.Add("pdfPath", new TextStructure(pdfPath));
            scripter.InitVariables.Add("xpsPath", new TextStructure(xpsPath));

            if (File.Exists(pdfPath))
                File.Delete(pdfPath);

            if (File.Exists(xpsPath))
                File.Delete(xpsPath);
        }

        [SetUp]
        public void TestInit()
        {
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.MSOffice.dll");
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelExportTest()
        {
            scripter.Text =($@"excel.open {SpecialChars.Variable}xlsPath
                              excel.export {SpecialChars.Variable}pdfPath
                              excel.export {SpecialChars.Variable}xpsPath");
            scripter.Run();
            Assert.IsTrue(File.Exists(xpsPath));
            Assert.IsTrue(File.Exists(pdfPath));
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout)]
        public void ExcelExportFailTest()
        {
            scripter.Text =($@"excel.open {SpecialChars.Variable}xlsPath
                              excel.export {SpecialChars.Text}C:\lol\ble.docx{SpecialChars.Text}");
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

            if (File.Exists(pdfPath))
                File.Delete(pdfPath);

            if (File.Exists(xpsPath))
                File.Delete(xpsPath);
        }
    }
}
