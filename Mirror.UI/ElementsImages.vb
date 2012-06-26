Imports Mirror.Core

Public Class ImageConverter
    Implements IValueConverter
    Dim Dic As New Dictionary(Of Element, ImageSource())

    ''' <summary>
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="parameter">It's System.Int32 value. 0 - element image; 1 - access image</param>
    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.Convert
        Dim p As Integer = parameter
        Dim e As Element = value
        Dim res = ImageResolver.ResolveElement(e)(p)
        If res Is Nothing Then Return New BitmapImage
        Return res
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As Globalization.CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New Exception
    End Function
End Class

Public Class ImageResolver
    Shared Sources As New Dictionary(Of String, ImageSource)

    Shared Function ResolveElement(e As Element) As ImageSource()
        Dim arr(1) As ImageSource
        arr(0) = LoadImage("Images\elements\" & e.ElementTypeString)
        If e.Access <> ElementAccess.Public Then
            arr(1) = LoadImage("Images\elements\" & e.AccessString)
        End If
        Return arr
    End Function

    Shared Function LoadImage(str As String) As ImageSource
        str = str & ".png"
        If Sources.ContainsKey(str) Then Return Sources(str)
        Dim s As BitmapImage = Nothing
        Try
            s = New BitmapImage
            s.BeginInit()
            s.UriSource = New Uri(str, UriKind.RelativeOrAbsolute)
            s.EndInit()
            Sources.Add(str, s)
        Catch ex As Exception
            Debugger.Break()
        End Try
        Return s
    End Function
End Class
