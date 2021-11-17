Imports System.Windows.Forms
Imports System.Drawing

Namespace Objects
    REM I need to confirm all this functions are working
    Public Class EArrays



        Public Shared Function valueOf(Of T)(pObj As Object) As T()
            If TypeOf pObj Is T() Then Return CType(pObj, T())
            Return Nothing
        End Function















        ''' <summary>
        ''' Join Arrays of the same types
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="arrys"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CombineArrays(Of T)(ByVal ParamArray arrys() As List(Of T)) As T()
            Dim lstArry As List(Of T) = New List(Of T)

            For Each arry As List(Of T) In arrys
                lstArry.AddRange(arry)
            Next

            Return lstArry.ToArray
        End Function


        ''' <summary>
        ''' Get Next Item in Array to CurrentItem. If current item=last item, return first element.
        ''' </summary>
        ''' <param name="strElements"></param>
        ''' <param name="CurrentItem">If current item is not found, returns first item</param>
        ''' <param name="Delimiter">Default is Comma(,)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetNextElementInArray(ByVal strElements As String,
                                              ByVal CurrentItem As String,
                                              Optional ByVal Delimiter As String = ","
                                              ) As String
            If strElements Is Nothing OrElse strElements = String.Empty Then Return String.Empty

            Dim strElementsArray As String() = Split(strElements, Delimiter)
            Dim indexOfCurrent As Integer = Array.IndexOf(strElementsArray, CurrentItem)
            If indexOfCurrent = -1 Then
                Return strElementsArray(0)
            ElseIf (indexOfCurrent + 1) <= (strElementsArray.Length - 1) Then
                Return strElementsArray(indexOfCurrent + 1)
            Else
                Return strElementsArray(0)
            End If
        End Function

#Region "Searching Array"

        ''' <summary>
        ''' Search array first column in 2 dimaensional array. Item is Exact
        ''' </summary>
        ''' <param name="strArray"></param>
        ''' <param name="strSearch"></param>
        ''' <param name="DirectCastEqualsTheValueNotInstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Search_in_Array(ByVal strArray(,) As String, ByVal strSearch As String,
                                    ByVal DirectCastEqualsTheValueNotInstr As Boolean) As Integer

            If strArray Is Nothing OrElse strSearch Is Nothing OrElse strSearch = String.Empty Then Return -1
            If Not DirectCastEqualsTheValueNotInstr Then Return Search_in_Array(strArray, strSearch)
            '
            '   Search element in one dimensional array and returns the index if found else returns -1
            '
            Dim elementIndex As Integer = -1
            Dim intCount As Integer = 0

            For intCount = 0 To strArray.GetUpperBound(0)

                If strArray(intCount, 0) = strSearch Then
                    elementIndex = intCount
                    Exit For
                End If

            Next

            Return elementIndex

        End Function

        ''' <summary>
        ''' Search array first column in 2 dimaensional array. Item Contains
        ''' </summary>
        ''' <param name="strArray"></param>
        ''' <param name="strSearch"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Search_in_Array(ByVal strArray(,) As String, ByVal strSearch As String) As Integer
            '
            '   Search element in one dimensional array and returns the index if found else returns -1
            '
            Dim elementIndex As Integer = -1
            Dim intCount As Integer = 0

            For intCount = 0 To strArray.GetUpperBound(0)

                If InStr(strArray(intCount, 0), strSearch, CompareMethod.Text) > 0 And
                    InStr(strArray(intCount, 0), strSearch & ",", CompareMethod.Text) = 0 Then
                    elementIndex = intCount
                    Exit For
                End If

            Next

            Return elementIndex

        End Function

        ''' <summary>
        ''' Indicates a Search in Array of Lines where the WHOLE Line must be equal to the String For Searching. It returns ONLY First INDEX
        ''' </summary>
        ''' <param name="strArray"></param>
        ''' <param name="strSearch"></param>
        ''' <param name="DirectCastEqualsTheValueNotInstr">Just Set to True, If you want this type of Search. 
        ''' Otherwise, Set to False Which Indicates ONLY part of the line is equal to the String For Searching</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Search_in_Array(ByVal strArray() As String, ByVal strSearch As String,
                                        ByVal DirectCastEqualsTheValueNotInstr As Boolean) As Integer

            If Not DirectCastEqualsTheValueNotInstr Then Return Search_in_Array(strArray, strSearch)
            '
            '   Search element in one dimensional array and returns the index if found else returns -1
            '
            Dim elementIndex As Integer = -1
            Dim intCount As Integer = 0

            For intCount = 0 To strArray.GetUpperBound(0)

                If strArray(intCount).Equals(strSearch) Then
                    elementIndex = intCount
                    Exit For
                End If

            Next

            Return elementIndex

        End Function

        ''' <summary>
        ''' Search through lines of array for a word string. If the word is contained in a line, it returns the line index [First Index]
        ''' </summary>
        ''' <param name="strArray">Lines (Of String) array</param>
        ''' <param name="strSearch">String to search</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Search_in_Array(ByVal strArray() As String, ByVal strSearch As String) As Integer
            '
            '   Search element in one dimensional array and returns the index if found else returns -1
            '
            Dim elementIndex As Integer = -1
            Dim intCount As Integer = 0

            ' Check through, If PResent
            '   Check also if extras has not been added Using comma(,) to detect
            '   InStr(strArray(intCount), strSearch & ",", CompareMethod.Text) = 0 
            For intCount = 0 To strArray.GetUpperBound(0)

                If InStr(strArray(intCount), strSearch, CompareMethod.Text) > 0 And
                    InStr(strArray(intCount), strSearch & ",", CompareMethod.Text) = 0 Then
                    elementIndex = intCount
                    Exit For
                End If

            Next

            Return elementIndex

        End Function

        ''' <summary>
        ''' Search through lines of array for a word string. If the word is contained in a line, it returns the line index. Select Which Index [First, Last ...]
        ''' </summary>
        ''' <param name="strArray">Lines (Of String) array</param>
        ''' <param name="strSearch">String to search</param>
        ''' <param name="IndexType">Indicates Which Index you want? If First Discovered or Last ....</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Search_in_Array(ByVal strArray() As String,
                                        ByVal strSearch As String,
                                        ByVal IndexType As SearchArrays
                                        ) As Integer

            'If first index indicated
            If IndexType = SearchArrays.FirstIndex Then
                Search_in_Array(
                                        strArray:=strArray,
                                        strSearch:=strSearch
                                        )
            End If


            '
            '   Search element in one dimensional array and returns the LastIndex if found else returns -1
            '
            Dim elementIndex As Integer = -1
            Dim intCount As Integer = 0

            ' Check through, If PResent
            '   Check also if extras has not been added Using comma(,) to detect
            Try
                '   InStr(strArray(intCount), strSearch & ",", CompareMethod.Text) = 0 
                For intCount = 0 To strArray.GetUpperBound(0)

                    If InStr(strArray(intCount), strSearch, CompareMethod.Text) > 0 And
                        InStr(strArray(intCount), strSearch & ",", CompareMethod.Text) = 0 Then
                        elementIndex = intCount
                        'Continue Looping Incase we find another one
                        'Exit For
                    End If

                Next
            Catch ex As Exception

            End Try

            Return elementIndex

        End Function

        ''' <summary>
        ''' Use to indicate which index user wants
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum SearchArrays
            FirstIndex
            LastIndex
        End Enum

#End Region



        ''' <summary>
        ''' Returns first column
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="TwoDimension"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetOneDimension(Of T)(ByVal TwoDimension As T(,)) As T()
            Dim OneDimension As T() = Nothing
            Try
                'Incase two dimension is null 
                TwoDimension.CopyTo(OneDimension, 0)

            Catch ex As Exception

            End Try
            Return OneDimension
        End Function




#Region "Disposing Object Collections"


        ''' <summary>
        ''' Dispose Objects in a array and sets the array to nothing. Throws Exception
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <remarks></remarks>
        Public Shared Sub Dispose_Objects_Collection(Of T As {IDisposable})(ByRef obj() As T)

            If obj Is Nothing Then Return

            For Each objChild As T In obj
                If Not IsNothing(objChild) Then
                    'RaiseEvent ProgressLoading(
                    '                obj.GetUpperBound(0),
                    '                Array.IndexOf(obj, objChild)
                    '                )

                    objChild.Dispose()
                End If
                Application.DoEvents()
            Next
            obj = Nothing

        End Sub






        ''' <summary>
        ''' Dispose Object and returns true on Success
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function IsObjectDisposedSuccessfully(obj As Control) As Boolean
            Try
                If obj IsNot Nothing AndAlso Not obj.IsDisposed AndAlso obj.IsHandleCreated Then obj.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function




        ''' <summary>
        ''' Dispose Objects in a array. Throws Exception.
        ''' NB: this doesn't clear the array.
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <remarks></remarks>
        Public Shared Function DisposeObjects(Of T As {Control})(ByRef obj() As T) As Boolean

            If obj Is Nothing Then Return True

            Dim dCount As Int32 = Aggregate ds As Control In obj
                             Where IsObjectDisposedSuccessfully(ds) AndAlso
                             ds.IsDisposed
                             Into Count()

            Return CBool(dCount)

        End Function








#End Region



#Region "Array Extensions"

        Public Shared Function LastOf(Of T)(pNumberOfElementToReturn As UInt32, pSource As IEnumerable(Of T)) As IEnumerable(Of T)
            If pNumberOfElementToReturn > pSource.Count() Then Throw New Exception("Number of elements to return is greater than the elements available")
            Dim pStartIndex As Int32 = CInt(pSource.Count() - pNumberOfElementToReturn)

            Dim rst As IEnumerable(Of T) = (From d As T In pSource
                                             Where pSource.ToList().IndexOf(d) >= pStartIndex
                                             Select d
                                             ).ToList()

            Return rst
        End Function

#End Region






        ''' <summary>
        ''' Get a copy of array .. Not pointing to same memory address
        ''' </summary>
        ''' <param name="arr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function getArrayCopy(Of T)(ByVal arr As T()) As T()
            Return arr.ToList.ToArray
        End Function


        Public Shared Sub SwapObjectIndexInArray(Of T As Control)(ByRef ArrayObjectCollection As T(), ByVal RealIndexOnArray As Integer, ByVal NewSuggestedIndex As Integer)
            Dim _List As List(Of T) = ArrayObjectCollection.ToList

            'This is a referenced declaration
            Dim ObjSwap As Control = _List(RealIndexOnArray)

            'Swap Properties
            Dim _Location As Point = ObjSwap.Location
            ObjSwap.Location = _List(NewSuggestedIndex).Location
            _List(NewSuggestedIndex).Location = _Location

            'Dont Change Tag. Else, you will have infinite loop
            'Tag is the key to the rearrangement
            ''Dim _Tag As String = ObjSwap.Tag
            ''ObjSwap.Tag = _List(NewSuggestedIndex).Tag
            ''_List(NewSuggestedIndex).Tag = _Tag

            'Swap Positions
            _List(RealIndexOnArray) = _List(NewSuggestedIndex)
            _List(NewSuggestedIndex) = CType(ObjSwap, T)

            'Return result
            ArrayObjectCollection = _List.ToArray
        End Sub

        ' ''' <summary>
        ' ''' Adds their index to their tag property
        ' ''' </summary>
        ' ''' <typeparam name="T"></typeparam>
        ' ''' <param name="ArrayObjectCollection"></param>
        ' ''' <remarks></remarks>
        'Public Shared Sub ReTagAllAllItemsAccordingToIndex(Of T As Control)(ByRef ArrayObjectCollection As T())
        '    Dim _List As List(Of T) = ArrayObjectCollection.ToList

        '    For Each _Item As T In ArrayObjectCollection

        '        _Item.Tag = Array.IndexOf(ArrayObjectCollection, _Item)

        '    Next
        'End Sub


        Public Shared Function convert_to_string(ByVal lst As List(Of String)) As String
            Return convert_to_string(lst, vbCrLf)
        End Function


        Public Shared Function convert_to_string(ByVal lst As List(Of String), ByVal Delimiter As String) As String
            Return convert_to_string(lst, Delimiter, String.Empty)
        End Function

        Public Shared Function convert_to_string(ByVal lst As List(Of String),
                                         ByVal Delimiter As String,
                                         ByVal PadFront As String) As String
            Return convert_to_string(lst, Delimiter, PadFront, False)
        End Function


        Public Shared Function convert_to_string(ByVal lst As List(Of String),
                                          ByVal Delimiter As String,
                                          ByVal PadFront As String,
                                          ByVal Numbered As Boolean) As String
            Dim result As String = String.Empty
            If Delimiter Is Nothing Then Delimiter = vbCrLf
            If lst Is Nothing Then Return result

            For i As Int16 = 0 To CShort(lst.Count - 1)

                If Numbered Then
                    result &= String.Format("{0}.) {1}{2}", i + 1, PadFront, lst(i))
                Else
                    result &= PadFront & lst(i)
                End If

                If i < lst.Count - 1 Then result &= Delimiter
            Next

            Return result
        End Function

    End Class
End Namespace