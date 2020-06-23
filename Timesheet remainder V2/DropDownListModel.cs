using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Timesheet_remainder
{
    internal static class DropDownListModel
    {
        private static readonly Queue<string> ListItems = new Queue<string>();

        static DropDownListModel()
        {
            for (int i = 0; i <= 5; i++)
            {
                AddToDropDownList(string.Empty);
            }

            //debug only
            /*for (int i = 0; i <= 6; i++)
            {
                AddToDropDownList($"item{i}");
            }*/
        }

        public static void AddToDropDownList(string itemToAdd)
        {
            ListItems.Enqueue(itemToAdd);

            if (ListItems.Count > 5)
            {
                ListItems.Dequeue();
            }
        }

        public static string GetItemAtIndex(int queueIndex)
        {
            return ListItems.ElementAt(queueIndex);
        }
    }
}
