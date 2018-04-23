using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace SimpleDataMerge
{
    /// <summary>
    /// Interaction logic for CustomizeMergeFieldsControl.xaml
    /// </summary>
    public partial class CustomizeMergeFieldsControl : UserControl
    {
        public CustomizeMergeFieldsControl(MergeFieldNameInfo[] mergeFieldsNames)
        {
            InitializeComponent();
            grid.ItemsSource = mergeFieldsNames;
        }
    }
}
