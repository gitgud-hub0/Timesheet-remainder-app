using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Timesheet_remainder
{
    internal class ComboBoxTaskInputController : ComboBox
    {
        private readonly ComboBox _comboBoxTaskInput;

        public ComboBoxTaskInputController(ComboBox comboBoxTaskInput)
        {
            _comboBoxTaskInput = comboBoxTaskInput;
        }

        public ComboBox PopulateDropDownBox()
        {
            int fixedListMaxIndex = 4;
            if (_comboBoxTaskInput.Items.Count != 0)
            {
                for (int i = 0; i <= fixedListMaxIndex; i++)
                {
                    _comboBoxTaskInput.Items.RemoveAt(_comboBoxTaskInput.Items.Count - 1);
                }

                for (int i = 0; i <= fixedListMaxIndex; i++)
                {
                    _comboBoxTaskInput.Items.Add(DropDownListModel.GetItemAtIndex(fixedListMaxIndex - i));
                }
            }
            else
            {
                for (int i = 0; i <= fixedListMaxIndex; i++)
                {
                    _comboBoxTaskInput.Items.Add(DropDownListModel.GetItemAtIndex(fixedListMaxIndex - i));
                }
            }

            return _comboBoxTaskInput;
        }
    }
}
