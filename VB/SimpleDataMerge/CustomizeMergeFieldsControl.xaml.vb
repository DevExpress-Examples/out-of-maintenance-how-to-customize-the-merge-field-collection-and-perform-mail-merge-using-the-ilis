Imports System.Windows.Controls

Namespace SimpleDataMerge

    ''' <summary>
    ''' Interaction logic for CustomizeMergeFieldsControl.xaml
    ''' </summary>
    Public Partial Class CustomizeMergeFieldsControl
        Inherits UserControl

        Public Sub New(ByVal mergeFieldsNames As MergeFieldNameInfo())
            Me.InitializeComponent()
            Me.grid.ItemsSource = mergeFieldsNames
        End Sub
    End Class
End Namespace
