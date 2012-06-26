Imports System.Reflection
Imports Mirror.Core.StringBuild

''' <summary>
''' Contains the whole information for each type of needed members
''' </summary>
Public Class Element

    <DebuggerBrowsable(0)> Dim _ElementType As AllTypes
    <DebuggerBrowsable(0)> Dim _Name As String
    <DebuggerBrowsable(0)> Dim builder As IStringBuilder
    <DebuggerBrowsable(0)> Dim _access As ElementAccess
    <DebuggerBrowsable(0)> Dim _ElementInstance As ElementInstance
    <DebuggerBrowsable(0)> Dim _IsReference As Boolean '' for assembly
    Dim _children As IEnumerable(Of Element)

    ''' <summary>
    ''' Contains specific data for different members
    ''' </summary>
    ''' <Namespace>
    ''' 00 - Types in namespace
    ''' </Namespace>
    ''' <remarks>
    ''' </remarks>
    Dim SpecialData(2) As Object

    Dim _Element As Object '' Object to be reflected (MemberInfo, Assembly...)
    Dim _AccessString As String
    Dim _ElementTypeString As String

    Sub New(e As MemberInfo, resolver As IStringBuilder)
        _Element = e
        Me.builder = resolver
        ReadInfo(e)
    End Sub

    Sub New(name As String, types As IEnumerable(Of Type), resolver As IStringBuilder)
        Me._Name = name
        Me.builder = resolver
        SpecialData(0) = types
        Me._ElementType = AllTypes.Namespace
    End Sub

    Sub New(e As Assembly, resolver As IStringBuilder)
        _Element = e
        Me.builder = resolver
        Me._ElementType = AllTypes.Assembly
        Me.ReadInfo(e)
    End Sub

    ReadOnly Property ElementType As AllTypes
        Get
            Return _ElementType
        End Get
    End Property

    ReadOnly Property Access As ElementAccess
        Get
            Return _access
        End Get
    End Property

    ReadOnly Property ElementInstance As ElementInstance
        Get
            Return _ElementInstance
        End Get
    End Property

    ReadOnly Property Name As String
        Get
            Return _Name
        End Get
    End Property

    ReadOnly Property Member As MemberInfo
        Get
            Return _Element
        End Get
    End Property

    ReadOnly Property IsReference As Boolean
        Get
            Return _IsReference
        End Get
    End Property

    ReadOnly Property [Namespace] As String
        Get
            If TypeOf _Element Is TypeInfo Then Return DirectCast(_Element, TypeInfo).Namespace
            If _ElementType = AllTypes.Namespace Then Return _Name
            If _ElementType <> AllTypes.Assembly Then Return DirectCast(_Element, MemberInfo).DeclaringType.Namespace
            Return String.Empty
        End Get
    End Property

#Region "Strings"
    ReadOnly Property ElementTypeString As String
        Get
            If _ElementTypeString Is Nothing Then
                _ElementTypeString = [Enum].GetName(GetType(AllTypes), _ElementType)
            End If
            Return _ElementTypeString
        End Get
    End Property

    ReadOnly Property AccessString As String
        Get
            If _AccessString Is Nothing Then
                _AccessString = [Enum].GetName(GetType(ElementAccess), _access)
            End If
            Return _AccessString
        End Get
    End Property

    ReadOnly Property ToolTip As String
        Get
            Dim flag = builder.UseFullNames
            If Not flag Then
                builder.UseFullNames = True
            End If
            ToolTip = FullTreeData
            builder.UseFullNames = flag
        End Get
    End Property

    Public ReadOnly Property ImageToolTip As String
        Get
            Return String.Join(" ", {GetAccessString(), GetInstanceString(), GetElementTypeString()}.Where(IsNotNullOrEmptyHandler))
        End Get
    End Property

    Private Shared IsNotNullOrEmptyHandler As New Func(Of String, Boolean)(AddressOf IsNotNullOrEmpty)
    Private Shared Function IsNotNullOrEmpty(v As String) As Boolean
        Return Not String.IsNullOrEmpty(v)
    End Function

    Private Function GetInstanceString() As String
        If _ElementType = AllTypes.Field Or _ElementType = AllTypes.Method Or _ElementType = AllTypes.Property Or _ElementType = AllTypes.Event Then
            If _ElementInstance = Core.ElementInstance.Static Then
                Return builder.StaticString '' 
            End If
        End If
        Return ""
    End Function

    Private Function GetElementTypeString() As String
        If _ElementType = AllTypes.EnumField Then Return "Field"
        Return ElementTypeString
    End Function

    Private Function GetAccessString() As String
        If _access = ElementAccess.ProtectedFriend Then Return "Protected Friend"
        Select Case _ElementType
            Case AllTypes.Assembly, AllTypes.EnumField, AllTypes.References, AllTypes.Namespace
                Return ""
        End Select
        Return AccessString
    End Function

    ReadOnly Property FullTreeData As String
        Get
            Select Case _ElementType
                Case AllTypes.Field
                    Return String.Format("{0}:{1}", Name, builder.BuildFieldType(DirectCast(_Element, FieldInfo)))
                Case AllTypes.Property
                    Return String.Format("{0}:{1}", Name, builder.BuildPropertyType(DirectCast(_Element, PropertyInfo)))
                Case AllTypes.Method
                    Return String.Format("{0}({1}):{2}", Name, builder.BuildMethodParamteres(DirectCast(_Element, MethodInfo)), builder.BuildMethodType(DirectCast(_Element, MethodInfo)))
                Case AllTypes.Constructor
                    Return String.Format("{0}({1})", Name, builder.BuildMethodParamteres(DirectCast(_Element, MethodBase)))
                Case AllTypes.Enum
                    Return String.Format("{0}:{1}", Name, builder.BuildType(DirectCast(_Element, TypeInfo).GetEnumUnderlyingType))
                Case AllTypes.Class, AllTypes.Interface, AllTypes.Structure, AllTypes.Enum, AllTypes.Module, AllTypes.Delegate
                    Return builder.BuildType(DirectCast(_Element, TypeInfo))
                Case Else
                    Return Name
            End Select
        End Get
    End Property
#End Region

    Private Sub ReadInfo(e As MemberInfo) '' the longest method
        If e.MemberType = MemberTypes.TypeInfo Then
            _Name = builder.BuildType(e)

        Else
            _Name = e.Name
        End If
        Select Case e.MemberType
            Case MemberTypes.Constructor
                Me._ElementType = AllTypes.Constructor
            Case MemberTypes.Method
                Me._ElementType = AllTypes.Method
            Case MemberTypes.Event
                Me._ElementType = AllTypes.Event
            Case MemberTypes.Property
                Me._ElementType = AllTypes.Property
            Case MemberTypes.Field
                If DirectCast(e, FieldInfo).DeclaringType.IsEnum AndAlso DirectCast(e, FieldInfo).Name <> "value__" Then '' this is only for visualization
                    _ElementType = AllTypes.EnumField
                Else
                    _ElementType = AllTypes.Field
                End If
            Case MemberTypes.TypeInfo, MemberTypes.NestedType
                Dim t As Type = e
                If t.BaseType = GetType(MulticastDelegate) Then
                    _ElementType = AllTypes.Delegate
                ElseIf t.IsInterface Then
                    _ElementType = AllTypes.Interface
                ElseIf t.IsEnum Then
                    _ElementType = AllTypes.Enum
                ElseIf t.IsValueType Then
                    If t.IsEnum Then
                        _ElementType = AllTypes.Enum
                    Else
                        _ElementType = AllTypes.Structure
                    End If
                Else
                    _ElementType = AllTypes.Class
                End If
                If Not t.IsPublic Or t.IsNestedAssembly Then
                    _access = ElementAccess.Friend
                ElseIf t.IsNestedPrivate Then
                    _access = ElementAccess.Private
                ElseIf t.IsNestedFamORAssem Then
                    _access = ElementAccess.Private
                End If
        End Select

        Select Case _ElementType
            Case AllTypes.Field
                Dim f As FieldInfo = _Element
                If f.IsPrivate Then Me._access = ElementAccess.Private : Exit Select
                If f.IsFamily Then Me._access = ElementAccess.Protected : Exit Select
                If f.IsStatic Then Me._ElementInstance = Core.ElementInstance.Static
            Case AllTypes.Method
                Dim m As MethodInfo = _Element
                ReadInfoFromMethodInfo(m)
            Case AllTypes.Property
                Dim f As PropertyInfo = _Element
                Dim m As MethodInfo
                If f.CanRead Then
                    m = f.GetGetMethod(True)
                Else
                    m = f.GetSetMethod(True)
                End If
                ReadInfoFromMethodInfo(m)
            Case AllTypes.Event
                Dim ev As EventInfo = _Element
                ReadInfoFromMethodInfo(ev.GetAddMethod())
        End Select
    End Sub

    Private Sub ReadInfoFromMethodInfo(m As MethodBase)
        If m Is Nothing Then Exit Sub
        If m.IsStatic Then Me._ElementInstance = Core.ElementInstance.Static
        If m.IsPrivate Then Me._access = ElementAccess.Private : Exit Sub
        If m.IsFamily Then Me._access = ElementAccess.Protected : Exit Sub
    End Sub

    Private Sub ReadInfo(e As Assembly)
        _Name = e.GetName.Name
    End Sub

#Region "Children"
    Private Function GetTypeChildren() As IEnumerable(Of Element)
        Return DirectCast(_Element, TypeInfo).GetMembers(BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static Or BindingFlags.Instance).Where(AddressOf ThisIsMyDirectChild).Select(AddressOf CreateElement)
    End Function

    'Private Function AddBaseTypeToChildrenIfNecessary(c As IEnumerable(Of Element)) As IEnumerable(Of Element)
    '    Dim basetype As Type = DirectCast(_Element, TypeInfo).BaseType
    '    If basetype Is Nothing Then
    '        Return c.Union({CreateElement( basetype})
    '    End If
    'End Function

    Private Function ThisIsMyDirectChild(m As MemberInfo) As Boolean
        Dim t As Type = _Element
        Return m.DeclaringType Is t
    End Function

    Private Function CreateElement(m As MemberInfo) As Element
        Return New Element(m, builder)
    End Function

    Private Function CreateElement(ass As AssemblyName) As Element
        Dim asss = New Element(Assembly.Load(ass.FullName), builder)
        asss._IsReference = True
        Return asss
    End Function

    Private Function IsNotSpecialName(e As Element) As Boolean
        If e._ElementType = AllTypes.Method Then
            Return Not DirectCast(e._Element, MethodInfo).IsSpecialName
        End If
        Return True
    End Function

    Private Function CreateElement(pair As KeyValuePair(Of String, List(Of Type))) As Element
        Return New Element(pair.Key, pair.Value, builder)
    End Function

    Private Shared Function CreateReferenceElement(ass As Assembly, resolver As IStringBuilder) As Element
        Return New Element(ass, resolver) With {._ElementType = AllTypes.References, ._Name = "References"}
    End Function

    Private Function GetAssemblyChildren() As IEnumerable(Of Element)
        Dim dic As New Dictionary(Of String, List(Of Type))
        For Each type In DirectCast(_Element, Assembly).GetTypes
            If type.IsNested Then Continue For
            Dim n = type.Namespace
            If String.IsNullOrEmpty(n) Then n = "-"
            If Not dic.ContainsKey(n) Then
                dic.Add(n, New List(Of Type))
            End If
            dic(n).Add(type)
        Next
        Dim ref = CreateReferenceElement(DirectCast(_Element, Assembly), builder)
        Return dic.Select(AddressOf CreateElement).Union({ref})
    End Function

    Private Function GetReferences() As IEnumerable(Of Element)
        Return DirectCast(_Element, Assembly).GetReferencedAssemblies.Select(AddressOf CreateElement)
    End Function

    ReadOnly Property HasChildren As Boolean
        Get
            Return _ElementType = AllTypes.Assembly OrElse TypeOf _Element Is TypeInfo OrElse _ElementType = AllTypes.Namespace OrElse _ElementType = AllTypes.References
        End Get
    End Property

    ReadOnly Property Element
        Get
            Return _Element
        End Get
    End Property

    ReadOnly Property Children As IEnumerable(Of Element)
        Get
            If _children Is Nothing Then
                If HasChildren Then
                    If Me._ElementType = AllTypes.Assembly Then
                        If _IsReference Then
                            _children = New Element() {}
                        Else
                            _children = GetAssemblyChildren()
                        End If
                    ElseIf _ElementType = AllTypes.Namespace Then
                        _children = DirectCast(SpecialData(0), IEnumerable(Of Type)).Select(AddressOf CreateElement).Where(AddressOf IsNotSpecialName).OrderBy(AddressOf OrderbyName).ToArray
                        Return _children
                    ElseIf TypeOf _Element Is TypeInfo Then
                        _children = GetTypeChildren()
                    ElseIf _ElementType = AllTypes.References Then
                        _children = GetReferences()
                    End If
                Else
                    _children = New Element() {}
                End If
            End If
            _children = _children.Where(AddressOf IsNotSpecialName).OrderBy(AddressOf OrderbyType).ThenBy(AddressOf OrderbyName).ToArray
            Return _children
        End Get
    End Property

    Function OrderbyName(e As Element) As String
        Return e._Name
    End Function

    Function OrderbyType(e As Element) As AllTypes
        Return e._ElementType
    End Function

#End Region
End Class


