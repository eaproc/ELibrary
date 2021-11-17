using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.VisualBasic;

namespace ELibrary.Standard.AppConfigurations
{
    public class CSVReader : IDisposable
    {



        #region Constructors

        public CSVReader(string csvFilePath, System.Text.Encoding encode, StringSplitOptions pSplitOptions, string RowDelimiters = Constants.vbCrLf, string CellDelimiter = ",")
        {
            if (!File.Exists(csvFilePath))
                return;
            try
            {
                string f = File.ReadAllText(csvFilePath, encode);
                if ((f.Trim() ?? "") == (string.Empty ?? ""))
                    return;

                // REM Parse File
                __RawRows = f.Split(new string[] { RowDelimiters }, StringSplitOptions.RemoveEmptyEntries).ToList();
                if (RawRows.Count > 0)
                {


                    // REM Now process the files using the real delimiter
                    foreach (string Line in RawRows)
                    {

                        // REM get keys values
                        var Cells = Line.Split(new string[] { CellDelimiter }, pSplitOptions);
                        __Rows.Add(Cells);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
            }

            __IsValid = true;
        }

        /// <summary>
        /// Uses ANSII - WINDOWS  Encoding.  System.Text.Encoding.Default
        /// </summary>
        /// <param name="csvFilePath"></param>
        /// <param name="RowDelimiters"></param>
        /// <param name="CellDelimiter"></param>
        /// <remarks></remarks>
        public CSVReader(string csvFilePath, StringSplitOptions pSplitOptions, string RowDelimiters = Constants.vbCrLf, string CellDelimiter = ",") : this(csvFilePath, System.Text.Encoding.Default, pSplitOptions, RowDelimiters, CellDelimiter)
        {
        }

        /// <summary>
        /// Uses ANSII - WINDOWS  Encoding.  System.Text.Encoding.Default. Removes empty cells 
        /// </summary>
        /// <param name="csvFilePath"></param>
        /// <param name="RowDelimiters"></param>
        /// <param name="CellDelimiter"></param>
        /// <remarks></remarks>
        public CSVReader(string csvFilePath, string RowDelimiters = Constants.vbCrLf, string CellDelimiter = ",") : this(csvFilePath, System.Text.Encoding.Default, StringSplitOptions.RemoveEmptyEntries, RowDelimiters, CellDelimiter)
        {
        }


        #endregion



        #region Properties

        private bool __IsValid = false;
        /// <summary>
        /// Checks only if file is valid. Use has rows to know if the file contains data
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool IsValid
        {
            get
            {
                return __IsValid;
            }
        }

        private List<string[]> __Rows = new List<string[]>();
        // 'Public ReadOnly Property Rows(Optional IgnoreFirstLine As Boolean = False) As List(Of String())
        // '    Get
        // '        If Me.hasRows AndAlso IgnoreFirstLine Then
        // '            If Me.__Rows.Count > 1 Then
        // '                Dim cpy As String()() = Array.CreateInstance(GetType(String()), Me.__Rows.Count - 1).Cast(Of String()).ToArray()
        // '                Me.__Rows.CopyTo(cpy, 1)
        // '                Return cpy.ToList()
        // '            End If

        // '            Return New List(Of String())
        // '        End If

        // '            Return Me.__Rows
        // '    End Get
        // 'End Property

        public List<string[]> Rows
        {
            get
            {
                return __Rows;
            }
        }

        private List<string> __RawRows = new List<string>();

        public List<string> RawRows
        {
            get
            {
                return __RawRows;
            }
        }

        public int Count
        {
            get
            {
                return __Rows.Count;
            }
        }

        public bool hasRows
        {
            get
            {
                return __Rows.Count > 0;
            }
        }


        #endregion




        #region IDisposable Support
        private bool disposedValue; // To detect redundant calls

        // IDisposable
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                    __Rows = null;
                    __RawRows = null;
                }

                // TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                // TODO: set large fields to null.
            }

            disposedValue = true;
        }

        public bool IsDisposed
        {
            get
            {
                return disposedValue;
            }
        }

        // TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        // Protected Overrides Sub Finalize()
        // ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        // Dispose(False)
        // MyBase.Finalize()
        // End Sub

        // This code added by Visual Basic to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion

    }
}