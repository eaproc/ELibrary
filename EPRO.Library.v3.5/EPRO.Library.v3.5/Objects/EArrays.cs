using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ELibrary.Standard.Objects
{
    // REM I need to confirm all this functions are working
    public class EArrays
    {
        public static T[] valueOf<T>(object pObj)
        {
            if (pObj is T[])
                return (T[])pObj;
            return null;
        }















        /// <summary>
        /// Join Arrays of the same types
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="arrys"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static T[] CombineArrays<T>(params List<T>[] arrys)
        {
            var lstArry = new List<T>();
            foreach (List<T> arry in arrys)
                lstArry.AddRange(arry);
            return lstArry.ToArray();
        }


        /// <summary>
        /// Get Next Item in Array to CurrentItem. If current item=last item, return first element.
        /// </summary>
        /// <param name="strElements"></param>
        /// <param name="CurrentItem">If current item is not found, returns first item</param>
        /// <param name="Delimiter">Default is Comma(,)</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string GetNextElementInArray(string strElements, string CurrentItem, string Delimiter = ",")
        {
            if (strElements is null || (strElements ?? "") == (string.Empty ?? ""))
                return string.Empty;
            var strElementsArray = Strings.Split(strElements, Delimiter);
            int indexOfCurrent = Array.IndexOf(strElementsArray, CurrentItem);
            if (indexOfCurrent == -1)
            {
                return strElementsArray[0];
            }
            else if (indexOfCurrent + 1 <= strElementsArray.Length - 1)
            {
                return strElementsArray[indexOfCurrent + 1];
            }
            else
            {
                return strElementsArray[0];
            }
        }

        #region Searching Array

        /// <summary>
        /// Search array first column in 2 dimaensional array. Item is Exact
        /// </summary>
        /// <param name="strArray"></param>
        /// <param name="strSearch"></param>
        /// <param name="DirectCastEqualsTheValueNotInstr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static int Search_in_Array(string[,] strArray, string strSearch, bool DirectCastEqualsTheValueNotInstr)
        {
            if (strArray is null || strSearch is null || (strSearch ?? "") == (string.Empty ?? ""))
                return -1;
            if (!DirectCastEqualsTheValueNotInstr)
                return Search_in_Array(strArray, strSearch);
            // 
            // Search element in one dimensional array and returns the index if found else returns -1
            // 
            int elementIndex = -1;
            int intCount = 0;
            var loopTo = strArray.GetUpperBound(0);
            for (intCount = 0; intCount <= loopTo; intCount++)
            {
                if ((strArray[intCount, 0] ?? "") == (strSearch ?? ""))
                {
                    elementIndex = intCount;
                    break;
                }
            }

            return elementIndex;
        }

        /// <summary>
        /// Search array first column in 2 dimaensional array. Item Contains
        /// </summary>
        /// <param name="strArray"></param>
        /// <param name="strSearch"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static int Search_in_Array(string[,] strArray, string strSearch)
        {
            // 
            // Search element in one dimensional array and returns the index if found else returns -1
            // 
            int elementIndex = -1;
            int intCount = 0;
            var loopTo = strArray.GetUpperBound(0);
            for (intCount = 0; intCount <= loopTo; intCount++)
            {
                if (Strings.InStr(strArray[intCount, 0], strSearch, CompareMethod.Text) > 0 & Strings.InStr(strArray[intCount, 0], strSearch + ",", CompareMethod.Text) == 0)
                {
                    elementIndex = intCount;
                    break;
                }
            }

            return elementIndex;
        }

        /// <summary>
        /// Indicates a Search in Array of Lines where the WHOLE Line must be equal to the String For Searching. It returns ONLY First INDEX
        /// </summary>
        /// <param name="strArray"></param>
        /// <param name="strSearch"></param>
        /// <param name="DirectCastEqualsTheValueNotInstr">Just Set to True, If you want this type of Search. 
        /// Otherwise, Set to False Which Indicates ONLY part of the line is equal to the String For Searching</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static int Search_in_Array(string[] strArray, string strSearch, bool DirectCastEqualsTheValueNotInstr)
        {
            if (!DirectCastEqualsTheValueNotInstr)
                return Search_in_Array(strArray, strSearch);
            // 
            // Search element in one dimensional array and returns the index if found else returns -1
            // 
            int elementIndex = -1;
            int intCount = 0;
            var loopTo = strArray.GetUpperBound(0);
            for (intCount = 0; intCount <= loopTo; intCount++)
            {
                if (strArray[intCount].Equals(strSearch))
                {
                    elementIndex = intCount;
                    break;
                }
            }

            return elementIndex;
        }

        /// <summary>
        /// Search through lines of array for a word string. If the word is contained in a line, it returns the line index [First Index]
        /// </summary>
        /// <param name="strArray">Lines (Of String) array</param>
        /// <param name="strSearch">String to search</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static int Search_in_Array(string[] strArray, string strSearch)
        {
            // 
            // Search element in one dimensional array and returns the index if found else returns -1
            // 
            int elementIndex = -1;
            int intCount = 0;

            // Check through, If PResent
            // Check also if extras has not been added Using comma(,) to detect
            // InStr(strArray(intCount), strSearch & ",", CompareMethod.Text) = 0 
            var loopTo = strArray.GetUpperBound(0);
            for (intCount = 0; intCount <= loopTo; intCount++)
            {
                if (Strings.InStr(strArray[intCount], strSearch, CompareMethod.Text) > 0 & Strings.InStr(strArray[intCount], strSearch + ",", CompareMethod.Text) == 0)
                {
                    elementIndex = intCount;
                    break;
                }
            }

            return elementIndex;
        }

        /// <summary>
        /// Search through lines of array for a word string. If the word is contained in a line, it returns the line index. Select Which Index [First, Last ...]
        /// </summary>
        /// <param name="strArray">Lines (Of String) array</param>
        /// <param name="strSearch">String to search</param>
        /// <param name="IndexType">Indicates Which Index you want? If First Discovered or Last ....</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static int Search_in_Array(string[] strArray, string strSearch, SearchArrays IndexType)
        {

            // If first index indicated
            if (IndexType == SearchArrays.FirstIndex)
            {
                Search_in_Array(strArray: strArray, strSearch: strSearch);
            }


            // 
            // Search element in one dimensional array and returns the LastIndex if found else returns -1
            // 
            int elementIndex = -1;
            int intCount = 0;

            // Check through, If PResent
            // Check also if extras has not been added Using comma(,) to detect
            try
            {
                // InStr(strArray(intCount), strSearch & ",", CompareMethod.Text) = 0 
                var loopTo = strArray.GetUpperBound(0);
                for (intCount = 0; intCount <= loopTo; intCount++)
                {
                    if (Strings.InStr(strArray[intCount], strSearch, CompareMethod.Text) > 0 & Strings.InStr(strArray[intCount], strSearch + ",", CompareMethod.Text) == 0)
                    {
                        elementIndex = intCount;
                        // Continue Looping Incase we find another one
                        // Exit For
                    }
                }
            }
            catch (Exception ex)
            {
            }

            return elementIndex;
        }

        /// <summary>
        /// Use to indicate which index user wants
        /// </summary>
        /// <remarks></remarks>
        public enum SearchArrays
        {
            FirstIndex,
            LastIndex
        }

        #endregion



        /// <summary>
        /// Returns first column
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="TwoDimension"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static T[] GetOneDimension<T>(T[,] TwoDimension)
        {
            T[] OneDimension = null;
            try
            {
                // Incase two dimension is null 
                TwoDimension.CopyTo(OneDimension, 0);
            }
            catch (Exception ex)
            {
            }

            return OneDimension;
        }




        #region Disposing Object Collections


        /// <summary>
        /// Dispose Objects in a array and sets the array to nothing. Throws Exception
        /// </summary>
        /// <param name="obj"></param>
        /// <remarks></remarks>
        public static void Dispose_Objects_Collection<T>(ref T[] obj) where T : IDisposable
        {
            if (obj is null)
                return;
            foreach (T objChild in obj)
            {
                if (!Information.IsNothing(objChild))
                {
                    // RaiseEvent ProgressLoading(
                    // obj.GetUpperBound(0),
                    // Array.IndexOf(obj, objChild)
                    // )

                    objChild.Dispose();
                }

                Application.DoEvents();
            }

            obj = null;
        }






        /// <summary>
        /// Dispose Object and returns true on Success
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private static bool IsObjectDisposedSuccessfully(Control obj)
        {
            try
            {
                if (obj is object && !obj.IsDisposed && obj.IsHandleCreated)
                    obj.Dispose();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }




        /// <summary>
        /// Dispose Objects in a array. Throws Exception.
        /// NB: this doesn't clear the array.
        /// </summary>
        /// <param name="obj"></param>
        /// <remarks></remarks>
        public static bool DisposeObjects<T>(ref T[] obj) where T : Control
        {
            if (obj is null)
                return true;
            int dCount = (from ds in obj
                          where IsObjectDisposedSuccessfully(ds) && ds.IsDisposed
                          select ds).Count();
            return Conversions.ToBoolean(dCount);
        }








        #endregion



        #region Array Extensions

        public static IEnumerable<T> LastOf<T>(uint pNumberOfElementToReturn, IEnumerable<T> pSource)
        {
            if (pNumberOfElementToReturn > pSource.Count())
                throw new Exception("Number of elements to return is greater than the elements available");
            int pStartIndex = (int)(pSource.Count() - pNumberOfElementToReturn);
            IEnumerable<T> rst = (from d in pSource
                                  where pSource.ToList().IndexOf(d) >= pStartIndex
                                  select d).ToList();
            return rst;
        }

        #endregion






        /// <summary>
        /// Get a copy of array .. Not pointing to same memory address
        /// </summary>
        /// <param name="arr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static T[] getArrayCopy<T>(T[] arr)
        {
            return arr.ToList().ToArray();
        }

        public static void SwapObjectIndexInArray<T>(ref T[] ArrayObjectCollection, int RealIndexOnArray, int NewSuggestedIndex) where T : Control
        {
            var _List = ArrayObjectCollection.ToList();

            // This is a referenced declaration
            Control ObjSwap = _List[RealIndexOnArray];

            // Swap Properties
            var _Location = ObjSwap.Location;
            ObjSwap.Location = _List[NewSuggestedIndex].Location;
            _List[NewSuggestedIndex].Location = _Location;

            // Dont Change Tag. Else, you will have infinite loop
            // Tag is the key to the rearrangement
            // 'Dim _Tag As String = ObjSwap.Tag
            // 'ObjSwap.Tag = _List(NewSuggestedIndex).Tag
            // '_List(NewSuggestedIndex).Tag = _Tag

            // Swap Positions
            _List[RealIndexOnArray] = _List[NewSuggestedIndex];
            _List[NewSuggestedIndex] = (T)ObjSwap;

            // Return result
            ArrayObjectCollection = _List.ToArray();
        }

        // ''' <summary>
        // ''' Adds their index to their tag property
        // ''' </summary>
        // ''' <typeparam name="T"></typeparam>
        // ''' <param name="ArrayObjectCollection"></param>
        // ''' <remarks></remarks>
        // Public Shared Sub ReTagAllAllItemsAccordingToIndex(Of T As Control)(ByRef ArrayObjectCollection As T())
        // Dim _List As List(Of T) = ArrayObjectCollection.ToList

        // For Each _Item As T In ArrayObjectCollection

        // _Item.Tag = Array.IndexOf(ArrayObjectCollection, _Item)

        // Next
        // End Sub


        public static string convert_to_string(List<string> lst)
        {
            return convert_to_string(lst, Constants.vbCrLf);
        }

        public static string convert_to_string(List<string> lst, string Delimiter)
        {
            return convert_to_string(lst, Delimiter, string.Empty);
        }

        public static string convert_to_string(List<string> lst, string Delimiter, string PadFront)
        {
            return convert_to_string(lst, Delimiter, PadFront, false);
        }

        public static string convert_to_string(List<string> lst, string Delimiter, string PadFront, bool Numbered)
        {
            string result = string.Empty;
            if (Delimiter is null)
                Delimiter = Constants.vbCrLf;
            if (lst is null)
                return result;
            for (short i = 0, loopTo = (short)(lst.Count - 1); i <= loopTo; i++)
            {
                if (Numbered)
                {
                    result += string.Format("{0}.) {1}{2}", i + 1, PadFront, lst[i]);
                }
                else
                {
                    result += PadFront + lst[i];
                }

                if (i < lst.Count - 1)
                    result += Delimiter;
            }

            return result;
        }
    }
}