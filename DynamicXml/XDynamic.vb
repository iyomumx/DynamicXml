Imports System.Dynamic
Imports System.Xml.Serialization
Imports System.Xml
Imports System.Xml.XPath

Public Class XDynamic
    Inherits DynamicObject
    Implements IXmlSerializable, IXmlLineInfo

    Protected Friend node As XElement
    Protected Friend nodes As List(Of XElement)

#Region "SubNew"
    ''' <summary>
    ''' 创建一个XDynamic的实例。
    ''' 请不要在代码里直接调用此过程，使用这个过程初始化的XDynamic对象不能正常工作，只作为标志使用。
    ''' 若需要一个空XDynamic对象，请使用XDynamicExtensions.EmptyXDynamic
    ''' </summary>
    ''' <remarks></remarks>
    Protected Friend Sub New()
        nodes = New List(Of XElement)
    End Sub

    ''' <summary>
    ''' 使用XDynamic包装特定的XElement对象，使之支持动态语句
    ''' </summary>
    ''' <param name="node">需要包装的XElement对象</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal node As XElement)
        Me.New()
        Me.node = node
        Me.nodes.Add(node)
    End Sub

    ''' <summary>
    ''' 使用XDynamic包装特定的XElement对象集合，使之支持动态语句
    ''' </summary>
    ''' <param name="nodes">需要包装的XElement对象集合</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal nodes As IEnumerable(Of XElement))
        Me.New()
        Me.nodes.AddRange(nodes)
        Me.node = nodes.FirstOrDefault
    End Sub

    ''' <summary>
    ''' 创建一个特定名称的空XElement并使用XDynamic包装
    ''' </summary>
    ''' <param name="name">XElement的名称</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal name As String)
        Me.New()
        Me.node = New XElement(name)
        Me.nodes.Add(node)
    End Sub

    ''' <summary>
    ''' 创建一个特定名称的空XElement并使用XDynamic包装
    ''' </summary>
    ''' <param name="name">XElement的名称</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal name As XName)
        Me.New()
        Me.node = New XElement(name)
        Me.nodes.Add(node)
    End Sub

#End Region

#Region "DynamicMethod"
    ''' <summary>
    ''' 尝试设置成员变量（此处为node的子级）
    ''' 若设置的变量不存在，则添加一个
    ''' 若设置值为XDynamic，则创建一个空项，否则将其设为所设值
    ''' </summary>
    ''' <param name="binder"></param>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function TrySetMember(ByVal binder As System.Dynamic.SetMemberBinder, ByVal value As Object) As Boolean
        Dim Xn = Me.node.Element(binder.Name)
        If Xn IsNot Nothing Then
            Xn.SetValue(value)
        Else
            If TypeOf value Is XDynamic Then
                Dim xd As XDynamic = value
                If xd.node IsNot Nothing Then
                    node.Add(xd.nodes.ToArray)
                Else
                    node.Add(New XElement(binder.Name))
                End If
            Else
                node.Add(New XElement(binder.Name, value))
            End If
        End If
        Return True
    End Function

    ''' <summary>
    ''' 按索引设置XElement集合的值。若指定的索引恰好等于集合项数，则创建一个XElement加入集合，并同时加入集合第一个元素的父节点内（若父节点存在）
    ''' </summary>
    ''' <param name="binder"></param>
    ''' <param name="indexes"></param>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function TrySetIndex(ByVal binder As System.Dynamic.SetIndexBinder, ByVal indexes() As Object, ByVal value As Object) As Boolean
        If indexes.Length = 0 Then
            Return False
        Else
            Try
                Dim Index = CInt(indexes(0))
                If Index = nodes.Count Then
                    Dim xe = If(CTypeDynamic(Of XElement)(value), New XDynamic(node.Name.ToString()))
                    If node.Parent IsNot Nothing Then
                        node.Parent.Add(xe)
                    End If
                    nodes.Add(xe)
                ElseIf Index > nodes.Count Then
                    Return False
                Else
                    nodes(Index) = value
                End If
            Catch ex As Exception
                Return False
            End Try
        End If
        Return True
    End Function

    ''' <summary>
    ''' 尝试对XDynamic对象进行索引。索引有三个重载：
    ''' <list>
    ''' <item>[int]：取指定位置的元素</item>
    ''' <item>[string]：对类中包装的XElement集合逐个计算指定的XPath，并返回IEnumerable类型的结果集合</item>
    ''' <item>[int,string]：对指定位置的元素计算指定的XPath，返回计算结果</item>
    ''' </list>
    ''' </summary>
    ''' <param name="binder"></param>
    ''' <param name="indexes"></param>
    ''' <param name="result"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function TryGetIndex(ByVal binder As System.Dynamic.GetIndexBinder, ByVal indexes() As Object, ByRef result As Object) As Boolean
        If indexes.Length = 0 Then
            result = Nothing
            Return False
        ElseIf indexes.Length = 1 Then
            Dim index = indexes(0)
            If TypeOf index Is Integer Then
                result = New XDynamic(nodes.ElementAt(index))
            ElseIf TypeOf index Is String Then
                result = nodes.Select(Function(nd) nd.XPathEvaluate(index))
            End If
        Else
            Dim index As Integer = indexes(0)
            Dim xpathstring As String = indexes(1)
            result = nodes(index).XPathEvaluate(xpathstring)
        End If
        Return True
    End Function

    ''' <summary>
    ''' 尝试获得XElement的子节点，并返回包装后的XDynamic对象
    ''' </summary>
    ''' <param name="binder"></param>
    ''' <param name="result"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function TryGetMember(ByVal binder As System.Dynamic.GetMemberBinder, ByRef result As Object) As Boolean
        Dim Xn = Me.node.Elements(binder.Name)
        If Xn IsNot Nothing Then
            result = New XDynamic(Xn)
            Return True
        Else
            result = Nothing
            Return False
        End If
    End Function

    ''' <summary>
    ''' 使用语言动态转换时调用，除了XElement支持的转换外，还支持转换为IEnumerable(Of XElement)以及IEnumerable(Of XDynamic)
    ''' </summary>
    ''' <param name="binder"></param>
    ''' <param name="result"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function TryConvert(ByVal binder As System.Dynamic.ConvertBinder, ByRef result As Object) As Boolean
        Select Case binder.Type
            Case GetType(IEnumerable(Of XElement))
                result = Me.nodes
                Return True
            Case GetType(IEnumerable(Of XDynamic))
                result = Me.nodes.Select(Function(n) New XDynamic(n))
                Return True
            Case Else
                Try
                    result = CTypeDynamic(node, binder.Type)
                    Return True
                Catch ex As Exception
                    result = Nothing
                    Return False
                End Try
        End Select
    End Function

    Protected Shared ReadOnly VBBindingGUID = Guid.Parse("6b7b62b2-9402-32eb-ba99-21e014026785")
    ''' <summary>
    ''' 尝试调用对象成员方法。此处添加对VB索引成员的支持。
    ''' </summary>
    ''' <param name="binder"></param>
    ''' <param name="args"></param>
    ''' <param name="result"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function TryInvokeMember(ByVal binder As System.Dynamic.InvokeMemberBinder, ByVal args() As Object, ByRef result As Object) As Boolean
        Try
            result = GetType(XElement).InvokeMember(binder.Name,
                                                    Reflection.BindingFlags.Public Or
                                                    Reflection.BindingFlags.Instance Or
                                                    Reflection.BindingFlags.InvokeMethod Or
                                                    If(binder.GetType.GUID = VBBindingGUID, Reflection.BindingFlags.GetProperty, 0), Nothing, node, args)
            Return True
        Catch ex As Exception
            If binder.GetType.GUID = VBBindingGUID Then
                Dim Xn = Me.node.Elements(binder.Name)
                If Xn IsNot Nothing Then
                    Dim index As Integer = -1
                    If args.Length <> 0 Then
                        Try
                            index = CInt(args(0))
                        Catch e As Exception
                            index = 0
                        End Try
                    Else
                        result = New XDynamic(Xn)
                        Return True
                    End If
                    If index = Xn.Count Then
                        Dim xe = New XElement(binder.Name)
                        node.Add(xe)
                        result = New XDynamic(xe)
                        Return True
                    End If
                    result = New XDynamic(Xn(index))
                    Return True
                End If
            End If
            result = Nothing
            Return True
        End Try
    End Function

    ''' <summary>
    ''' 对动态语言提供删除对象支持（测试中）
    ''' </summary>
    ''' <param name="binder"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function TryDeleteMember(ByVal binder As System.Dynamic.DeleteMemberBinder) As Boolean
        Dim Xn = node.Element(binder.Name)
        If Xn IsNot Nothing Then
            Xn.Remove()
            Return True
        Else
            Return False
        End If
    End Function

#End Region

#Region "Overrides Object Method"
    ''' <summary>
    ''' 返回节点列表的XML
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function ToString() As String
        Dim sb = New Text.StringBuilder
        For Each xe In nodes
            sb.AppendLine(xe.ToString)
        Next
        Return sb.ToString(0, If(sb.Length <> 0, sb.Length - Environment.NewLine.Length, 0))
    End Function
    ''' <summary>
    ''' 返回节点列表的XML，还可以选择禁用格式设置
    ''' </summary>
    ''' <param name="options">一个指定格式设置行为的System.Xml.Linq.SaveOptions</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function ToString(ByVal options As SaveOptions) As String
        Dim sb = New Text.StringBuilder
        For Each xe In nodes
            sb.AppendLine(xe.ToString(options))
        Next
        Return sb.ToString
    End Function
    ''' <summary>
    ''' 确定指定的Object是否等于当前的object
    ''' </summary>
    ''' <param name="obj">需要比较的Object</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Return nodes.Equals(obj)
    End Function

    ''' <summary>
    ''' 用作特定类型的哈希函数
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function GetHashCode() As Integer
        Return nodes.GetHashCode()
    End Function

#End Region

#Region "IXmlSerializable Member"
    Protected Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
        Return CType(node, IXmlSerializable).GetSchema()
    End Function

    Protected Sub ReadXml(ByVal reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml
        CType(node, IXmlSerializable).ReadXml(reader)
    End Sub

    Protected Sub WriteXml(ByVal writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml
        CType(node, IXmlSerializable).WriteXml(writer)
    End Sub
#End Region

#Region "IXmlLineInfo Member"

    Protected Function HasLineInfo() As Boolean Implements System.Xml.IXmlLineInfo.HasLineInfo
        Return CType(node, IXmlLineInfo).HasLineInfo()
    End Function

    <XmlIgnore()>
    Protected ReadOnly Property LineNumber As Integer Implements System.Xml.IXmlLineInfo.LineNumber
        Get
            Return CType(node, IXmlLineInfo).LineNumber
        End Get
    End Property

    <XmlIgnore()>
    Protected ReadOnly Property LinePosition As Integer Implements System.Xml.IXmlLineInfo.LinePosition
        Get
            Return CType(node, IXmlLineInfo).LinePosition
        End Get
    End Property

#End Region

#Region "Convert"

#Region "Narrowing(Port from XElement)"

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As String
        Return CType(obj.node, String)
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Boolean
        Return CType(obj.node, Boolean)
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Nullable(Of Boolean)
        Return CType(obj.node, Nullable(Of Boolean))
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Double
        Return CType(obj.node, Double)
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Nullable(Of Double)
        Return CType(obj.node, Nullable(Of Double))
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Single
        Return CType(obj.node, Single)
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Nullable(Of Single)
        Return CType(obj.node, Nullable(Of Single))
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Int32
        Return CType(obj.node, Int32)
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Nullable(Of Int32)
        Return CType(obj.node, Nullable(Of Int32))
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Int64
        Return CType(obj.node, Int64)
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Nullable(Of Int64)
        Return CType(obj.node, Nullable(Of Int64))
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As UInt32
        Return CType(obj.node, UInt32)
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Nullable(Of UInt32)
        Return CType(obj.node, Nullable(Of UInt32))
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Decimal
        Return CType(obj.node, Decimal)
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Nullable(Of Decimal)
        Return CType(obj.node, Nullable(Of Decimal))
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As DateTime
        Return CType(obj.node, DateTime)
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Nullable(Of DateTime)
        Return CType(obj.node, Nullable(Of DateTime))
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Guid
        Return CType(obj.node, Guid)
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Nullable(Of Guid)
        Return CType(obj.node, Nullable(Of Guid))
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As DateTimeOffset
        Return CType(obj.node, DateTimeOffset)
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Nullable(Of DateTimeOffset)
        Return CType(obj.node, Nullable(Of DateTimeOffset))
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As TimeSpan
        Return CType(obj.node, TimeSpan)
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As Nullable(Of TimeSpan)
        Return CType(obj.node, Nullable(Of TimeSpan))
    End Operator

#End Region

#Region "Convert To XElement or IEnumerable (For LINQ)"
    Public Shared Widening Operator CType(ByVal obj As XDynamic) As XElement
        Return obj.node
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As List(Of XDynamic)
        Return obj.nodes.Select(Function(node) New XDynamic(node)).ToList
    End Operator

    Public Shared Narrowing Operator CType(ByVal obj As XDynamic) As List(Of XElement)
        Return obj.nodes
    End Operator

    Public Shared Widening Operator CType(ByVal obj As XDynamic) As Object()
        Return obj.nodes.Select(Function(node) New XDynamic(node)).ToArray
    End Operator

#End Region

#End Region

End Class
