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
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice
{
    public static class ExcelManager
    {
        private static List<ExcelWrapper> launchedExcels = new List<ExcelWrapper>();

        private static ExcelWrapper currentExcel = null;

        public static ExcelWrapper CurrentExcel
        {
            get
            {
                if (currentExcel == null)
                {
                    throw new ApplicationException("Excel instance must be opened first using excel.open command");
                }
                return currentExcel;
            }
            private set
            {
                currentExcel = value;
            }
        }

        public static void SwitchExcel(int id)
        {
            ExcelWrapper instanceToSwitchTo = launchedExcels.Where(x => x.Id == id).FirstOrDefault();
            if (instanceToSwitchTo == null)
            {
                throw new ArgumentException($"No excel instance found with id: {id}");
            }
            CurrentExcel = instanceToSwitchTo;
        }

        private static int GetNextId()
        {
            return launchedExcels.Count() > 0 ? launchedExcels.Max(x => x.Id) + 1 : 0;
        }

        public static ExcelWrapper CreateInstance()
        {
            int assignedId = GetNextId();
            ExcelWrapper wrapper = new ExcelWrapper(assignedId);
            launchedExcels.Add(wrapper);
            CurrentExcel = wrapper;
            return wrapper;
        }

        public static void RemoveInstance(int? id = null)
        {
            if (id == null)
            {
                id = CurrentExcel.Id;
            }
            var toRemove = launchedExcels.Where(x => x.Id == id).FirstOrDefault();
            if (toRemove != null)
            {
                launchedExcels.Remove(toRemove);
                toRemove.Close();
            }
            else
            {
                throw new ArgumentException($"Unable to close excel instance with specified id argument: '{id}'");
            }            
        }
    }
}
