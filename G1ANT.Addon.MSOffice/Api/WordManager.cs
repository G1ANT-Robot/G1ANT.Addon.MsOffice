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
    public static class WordManager
	{
		private static List<WordWrapper> launchedWords = new List<WordWrapper>();

		public static WordWrapper CurrentWord { get; private set; }

		public static WordWrapper AddWord()
		{
			WordWrapper wrapper = new WordWrapper();
			launchedWords.Add(wrapper);
			CurrentWord = wrapper;
			return wrapper;
		}

		internal static int GetNextId()
		{
			return launchedWords.Count() > 0 ? launchedWords.Max(x => x.Id) + 1 : 0;
		}

		internal static bool Switch(int id)
		{
			WordWrapper ww = launchedWords.Where(x => x.Id == id).FirstOrDefault();
			CurrentWord = ww ?? CurrentWord;
			CurrentWord.Show();
			return ww != null;
		}

		public static void Remove(WordWrapper wordWrapper)
		{
			launchedWords.Remove(wordWrapper);
		}
	}
}
