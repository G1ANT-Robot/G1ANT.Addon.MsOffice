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
using System.Text;
using System.Threading.Tasks;

namespace G1ANT.Addon.MSOffice
{
    public static class OutlookManager
    {

        private static List<OutlookWrapper> launchedOutlooks = new List<OutlookWrapper>();

        private static OutlookWrapper currentOutlook;

        public static OutlookWrapper CurrentOutlook
        {
            get
            {
                if (currentOutlook == null)
                {
                    throw new ApplicationException("Outlook instance must be opened first using outlook.open command");
                }
                return currentOutlook;
            }
            set { currentOutlook = value; }
        }

        public static bool SwitchOutlook(int id)
        {
            var tmpOutlook = launchedOutlooks.Where(x => x.Id == id).FirstOrDefault();
            CurrentOutlook = tmpOutlook ?? CurrentOutlook;
            return tmpOutlook != null;
        }

        private static int GetNextId()
        {
            return launchedOutlooks.Count() > 0 ? launchedOutlooks.Max(x => x.Id) + 1 : 0;

        }

        public static OutlookWrapper AddOutlook()
        {
            int assignedId = GetNextId();
            OutlookWrapper wrapper = new OutlookWrapper(assignedId);
            launchedOutlooks.Add(wrapper);
            CurrentOutlook = wrapper;
            return wrapper;

        }

        public static void Remove(int id)
        {
            var toRemove = launchedOutlooks.Where(x => x.Id == id).FirstOrDefault();
            launchedOutlooks.Remove(toRemove);
        }

        public static string RemoveLineEndings(this string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return value;
            }
            string lineSeparator = ((char)0x2028).ToString();
            string paragraphSeparator = ((char)0x2029).ToString();

            return value.Replace("\r\n", string.Empty).Replace("\n", string.Empty).Replace("\r", string.Empty).Replace(lineSeparator, string.Empty).Replace(paragraphSeparator, string.Empty);
        }

    }
}
