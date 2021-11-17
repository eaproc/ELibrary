

Namespace Modules


    Public Module basReflection


        <System.Runtime.CompilerServices.Extension()> _
        Public Function getAnonymousClassPropValue(pClass As Object, pPropertyName As String,
                                                            Optional ByVal pIndex() As Object = Nothing) As Object

            If pClass Is Nothing Then Return Nothing

            Dim myType As Type = pClass.GetType()
            Dim props As IList(Of System.Reflection.PropertyInfo) = myType.GetProperties()
            Dim fields As IList(Of System.Reflection.FieldInfo) = myType.GetFields()
            If props.Count > 0 Then
                Dim prop = props.Where(Function(x) x.Name = pPropertyName).FirstOrDefault()

                If prop IsNot Nothing Then Return prop.GetValue(pClass, pIndex)
            ElseIf fields.Count > 0 Then
                Dim field = fields.Where(Function(x) x.Name = pPropertyName).FirstOrDefault()

                If field IsNot Nothing Then Return field.GetValue(pClass)

            End If

            Return Nothing

        End Function



        <System.Runtime.CompilerServices.Extension()> _
        Public Function getAnonymousClassMethod(pClass As Object, pMethodName As String) As Reflection.MethodInfo

            If pClass Is Nothing Then Return Nothing

            Dim myType As Type = pClass.GetType()
            Dim meths As IList(Of System.Reflection.MethodInfo) = myType.GetMethods()
            If meths.Count > 0 Then
                Dim meth = meths.Where(Function(x) x.Name = pMethodName).FirstOrDefault()
                If meth IsNot Nothing Then Return meth
            End If

            Return Nothing

        End Function




    End Module


End Namespace