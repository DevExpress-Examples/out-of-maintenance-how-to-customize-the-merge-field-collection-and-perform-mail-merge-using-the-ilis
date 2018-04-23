Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Windows
Imports System.Windows.Controls

Namespace SimpleDataMerge
    ''' <summary>
    ''' Interaction logic for CustomizeMergeFieldsControl.xaml
    ''' </summary>
    Partial Public Class CustomizeMergeFieldsControl
        Inherits UserControl

        Public Sub New(ByVal mergeFieldsNames() As MergeFieldNameInfo)
            InitializeComponent()
            grid.ItemsSource = mergeFieldsNames
        End Sub
    End Class
End Namespace
