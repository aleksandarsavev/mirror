Public Class AssembliesTree
    Inherits TreeView

    Shared dic As New MainDictionary

    Private Sub AssembliesTree_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        'Me.ItemsPanel = dic.Item("VSP")
        Me.ItemTemplate = dic.Item("ElementTampl")
        Me.ItemContainerStyle = dic.Item("TreeItemStyle")
    End Sub

    Private Sub AssembliesTree_SelectedItemChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles Me.SelectedItemChanged

    End Sub
End Class

Public Class LinkedBlock
    Inherits TextBlock

    Private Sub LinkedBlock_MouseEnter(sender As Object, e As MouseEventArgs) Handles Me.MouseEnter
        Me.TextDecorations.Add(Windows.TextDecorations.Underline)
    End Sub

    Private Sub LinkedBlock_MouseLeave(sender As Object, e As MouseEventArgs) Handles Me.MouseLeave
        Me.TextDecorations.Clear() ''(Windows.TextDecorations.Underline)
    End Sub
End Class
