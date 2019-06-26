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
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;

namespace G1ANT.Addon.MSOffice.Tests
{
    [TestFixture]
    [Apartment(ApartmentState.STA)]
    public class OutlookCloseTests
	{
        Scripter scripter;
        private void KillProcesses()
        {
            foreach (Process p in Process.GetProcessesByName("outlook"))
            {
                try
                {
                    p.Kill();
                }
                catch { }
            }
        }


        [OneTimeSetUp]
        public static void ClassInit()
        {            
            Environment.CurrentDirectory = TestContext.CurrentContext.TestDirectory;           
        }

        [SetUp]
        public void SetUp()
        {
            Language.Addon addon = Language.Addon.Load(@"G1ANT.Addon.MSOffice.dll");
            scripter = new Scripter();
scripter.InitVariables.Clear();
        }

        [Test]
        [Timeout(MSOfficeTests.TestsTimeout + 30000)]
		public void OutlookCloseTest()
		{
            KillProcesses();
            //System.Threading.Thread.Sleep(5000);
            Process[] userProcesses = Process.GetProcessesByName("outlook");
            scripter.RunLine("outlook.open");
            int tick = 0;
            int starttick = Environment.TickCount;
            int openingDelay = 50000;
            Process[] allProcesses = Process.GetProcessesByName("outlook");

            while (allProcesses.Length <= userProcesses.Length && tick <= starttick + openingDelay)
            {
                allProcesses = Process.GetProcessesByName("outlook");
                tick = Environment.TickCount;
                Thread.Sleep(10);
            }
            //int beforeCount = Process.GetProcessesByName("outlook").Length;
			//scripter.RunLine("outlook.open");
            //System.Threading.Thread.Sleep(5000);
            //int runCount = Process.GetProcessesByName("outlook").Length;
            //Assert.IsTrue(runCount > beforeCount);
            Assert.IsTrue(allProcesses.Length > userProcesses.Length);


            List<Process> diffProcesses = new List<Process>();

            foreach (var proc in allProcesses)
            {
                if (!userProcesses.Contains(proc))
                    diffProcesses.Add(proc);
            }
            scripter.RunLine("outlook.close");
            //System.Threading.Thread.Sleep(25000);
            //int closeCount = Process.GetProcessesByName("outlook").Length;
            allProcesses = Process.GetProcessesByName("outlook");
            Assert.AreEqual(userProcesses.Length, allProcesses.Length);
            //Assert.AreEqual(beforeCount, closeCount);
		}

        [TearDown]
        public void TestCleanUp()
        {
            Process[] proc = Process.GetProcessesByName("outlook");
            if (proc.Length != 0)
            {
                KillProcesses();
            }
        }
    }
}
