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

    Protected Friend Sub New()
        nodes = New List(Of XElement)
    End Sub

    Public Sub New(ByVal node As XElement)
        Me.New()
        Me.node = node
        Me.nodes.Add(node)
    End Sub

    Public Sub New(ByVal nodes As IEnumerable(Of XElement))
        Me.New()
        Me.nodes.AddRange(nodes)
        Me.node = nodes.FirstOrDefault
    End Sub

    Public Sub New(ByVal name As String)
        Me.New()
        Me.node = New XElement(name)
        Me.nodes.Add(node)
    End Sub

#End Region

#Region "DynamicMethod"

    Public Overrides Function TrySetMember(ByVal binder As System.Dynamic.SetMemberBinder, ByVal value As Object) As Boolean
        Dim Xn = Me.node.Element(binder.Name)
        If Xn IsNot Nothing Then
            Xn.SetValue(value)
        Else
            If TypeOf value Is XDynamic Then
                node.Add(New XElement(binder.Name))
            Else
                node.Add(New XElement(binder.Name, value))
            End If
        End If
        Return True
    End Function

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
                    Dim index As Integer
                    If args.Length <> 0 Then
                        Try
                            index = CInt(args(0))
                        Catch e As Exception
                            index = 0
                        End Try
                    Else
                        result = If(Xn.Count = 1,
                                    New XDynamic(Xn(0)),
                                    Xn.Select(Function(x) New XDynamic(x)))
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

    Public Overrides Function ToString() As String
        Dim sb = New Text.StringBuilder
        For Each xe In nodes
            sb.AppendLine(xe.ToString)
        Next
        Return sb.ToString
    End Function

    Public Overloads Function ToString(ByVal options As SaveOptions) As String
        Dim sb = New Text.StringBuilder
        For Each xe In nodes
            sb.AppendLine(xe.ToString(options))
        Next
        Return sb.ToString
    End Function

    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Return nodes.Equals(obj)
    End Function

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
