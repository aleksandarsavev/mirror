Imports System.Reflection
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace StringBuild
    Public Class VBStringBuilder
        Implements IStringBuilder

        Dim Resolved As New Dictionary(Of MemberInfo, String())

        Sub New()
            _UseFullNames = False
        End Sub

        Public Function BuildType(type As Type) As String Implements IStringBuilder.BuildType
            If type Is Nothing Then Throw New ArgumentException("Type cannot be nothing!", "type")
            If type.FullName Is Nothing Then Return type.Name '' generic argument type

            If Not type.IsGenericType AndAlso Resolved.ContainsKey(type) Then
                Dim res = Resolved(type)
                If UseFullNames Then
                    Return res(0) & res(1)
                Else
                    Return res(1)
                End If
            End If

            Dim nameS As String = ""
            nameS = type.Namespace
            If nameS <> "" Then nameS &= "."
            If type.IsNested Then
                nameS = BuildType(type.ReflectedType) & "." '' build full path of nasted type
            End If

            Dim name As String = ""

            If type.IsGenericType Then
                name = name & GenericBuilder(type, type.GetGenericArguments)
            Else
                name = name & type.Name
                If type.IsArray Then
                    name = name.Replace("[", "(").Replace("]", ")")

                End If
            End If
            If Not type.IsGenericType Then Resolved.Add(type, {nameS, name})
            Return name
        End Function

        Public Function BuildMethodName(m As MethodInfo) As String Implements IStringBuilder.BuildMethodName
            If Not m.IsGenericMethod Then Return m.Name
            Return GenericBuilder(m, m.GetGenericArguments)
        End Function

        Private Function GenericBuilder(info As MemberInfo, genTypes As Type()) As String
            Dim name As String = info.Name.Split("`"c).First ''removes the ugly sign and count of gen. arguments
            Dim genArgs = String.Join(", ", genTypes.Select(AddressOf BuildType))
            Return String.Format("{0}(Of {1})", name, genArgs)
        End Function

        Public Function BuildFieldType(f As FieldInfo) As String Implements IStringBuilder.BuildFieldType
            Return BuildType(f.FieldType)
        End Function

        Public Function BuildMethodParamteres(m As MethodBase) As String Implements IStringBuilder.BuildMethodParamteres
            Return String.Join(", ", m.GetParameters.Select(AddressOf BuildParameterInfo))
        End Function

        Private Function BuildParameterInfo(p As ParameterInfo) As String
            Dim coll As New List(Of String)
            If p.IsOptional Then coll.Add("Optional")
            If p.IsIn Then coll.Add("ByVal")
            If p.IsOut Then coll.Add("ByRef")
            coll.Add(p.Name)
            coll.Add("As")
            coll.Add(BuildType(p.ParameterType))
            Return String.Join(" ", coll)
        End Function

        Public Function BuildPropertyType(f As PropertyInfo) As String Implements IStringBuilder.BuildPropertyType
            Return BuildType(f.PropertyType)
        End Function

        Public Function BuildMethodType(m As MethodInfo) As String Implements IStringBuilder.BuildMethodType
            Return BuildType(m.ReturnType)
        End Function

        Public Property UseFullNames As Boolean Implements IStringBuilder.UseFullNames

        Public ReadOnly Property StaticString As String Implements IStringBuilder.StaticString
            Get
                Return "Shared"
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Creates logic for building of the strings
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface IStringBuilder

        ReadOnly Property StaticString As String

        Function BuildType(t As Type) As String
        Function BuildMethodName(m As MethodInfo) As String
        Function BuildMethodType(m As MethodInfo) As String
        Function BuildMethodParamteres(m As MethodBase) As String
        Function BuildFieldType(f As FieldInfo) As String
        Function BuildPropertyType(f As PropertyInfo) As String

        Property UseFullNames As Boolean
    End Interface

End Namespace