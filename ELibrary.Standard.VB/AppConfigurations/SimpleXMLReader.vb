
Namespace AppConfigurations

    ''' <summary>
    ''' Purpose of this class is to read an XML value without error. Regardless of the arrangement of the elements.
    ''' Please note that XML tag are case sensitive
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SimpleXMLReader



        Public Sub New(pXMLContent As String)

            Me.______ParsedXMLFile = XDocument.Parse(pXMLContent)

        End Sub








        Private ______ParsedXMLFile As XDocument
        Public ReadOnly Property ParsedXMLFile As XDocument
            Get
                Return Me.______ParsedXMLFile
            End Get
        End Property




        Public ReadOnly Property IsValid As Boolean
            Get
                Return Me.______ParsedXMLFile IsNot Nothing
            End Get
        End Property





        ''' <summary>
        ''' Returns all occurrences of this TagElements
        ''' </summary>
        ''' <param name="pTagName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function getElementsByTagName(pTagName As String) As IEnumerable(Of XElement)
            If Not Me.IsValid Then Return Nothing
            Dim p = Me.______ParsedXMLFile.Descendants(pTagName).ToList()

            Return p

        End Function


        ''' <summary>
        ''' get first occurence of this element
        ''' </summary>
        ''' <param name="pTagName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function getElementByTagName(pTagName As String) As XElement
            Dim p = getElementsByTagName(pTagName)
            If p IsNot Nothing Then Return p.First()
            Return Nothing
        End Function



        Public Function getElementsByTagNameValues(pTagName As String) As IEnumerable(Of String)
            Dim p = getElementsByTagName(pTagName)

            If p IsNot Nothing Then Return p.Select(Function(x) x.Value).ToList()
            Return Nothing


        End Function



        Public Function getElementByTagNameValue(pTagName As String) As String
            Dim p = getElementByTagName(pTagName)
            If p IsNot Nothing Then Return p.Value
            Return String.Empty

        End Function








    End Class


End Namespace