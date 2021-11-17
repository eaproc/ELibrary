using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CODERiT.Logger.v._3._5.Exceptions;
using Microsoft.VisualBasic;

namespace ELibrary.Standard.AppConfigurations
{
    public class IniReader
    {

        #region Constructors

        /// <summary>
        /// Reads Ini File.
        /// </summary>
        /// <param name="iniFilePath">The Ini File to read</param>
        /// <param name="KeyVDelimiter">The key value delimiter like Equals sign [ = ]</param>
        /// <param name="LineVDelimiter">The Line delimiter or Entry delimiter like New Line [ \n ]</param>
        /// <param name="pisCaseSensitive">Indicate if fetching result will be case sensitive</param>
        /// <remarks></remarks>
        public IniReader(string iniFilePath, System.Text.Encoding encode, string KeyVDelimiter, string LineVDelimiter, bool pisCaseSensitive, bool pIdentifySections) : this(iniFilePath, encode, KeyVDelimiter, LineVDelimiter, pisCaseSensitive, pIdentifySections, new string[] { })
        {
        }

        /// <summary>
        /// Reads Ini File. Uses ANSII - WINDOWS  Encoding.  System.Text.Encoding.Default
        /// </summary>
        /// <param name="iniFilePath">The Ini File to read</param>
        /// <param name="KeyVDelimiter">The key value delimiter like Equals sign [ = ]</param>
        /// <param name="LineVDelimiter">The Line delimiter or Entry delimiter like New Line [ \n ]</param>
        /// <param name="isCaseSensitive">Indicate if fetching result will be case sensitive</param>
        /// <remarks></remarks>
        public IniReader(string iniFilePath, string KeyVDelimiter, string LineVDelimiter, bool isCaseSensitive, bool pIdentifySections) : this(iniFilePath, System.Text.Encoding.Default, KeyVDelimiter, LineVDelimiter, isCaseSensitive, pIdentifySections, new string[] { })
        {
        }

        /// <summary>
        /// Reads Ini File.
        /// </summary>
        /// <param name="iniFilePath">The Ini File to read</param>
        /// <param name="isCaseSensitive">Indicate if fetching result will be case sensitive</param>
        /// <remarks></remarks>
        public IniReader(string iniFilePath, System.Text.Encoding encode, bool isCaseSensitive, bool pIdentifySections) : this(iniFilePath, encode, DEFAULT_KeyValuePairDelimiter, DEFAULT_LineDelimiter, isCaseSensitive, pIdentifySections, new string[] { })
        {
        }

        /// <summary>
        /// Reads Ini File. Uses ANSII - WINDOWS  Encoding.  System.Text.Encoding.Default
        /// </summary>
        /// <param name="iniFilePath">The Ini File to read</param>
        /// <param name="isCaseSensitive">Indicate if fetching result will be case sensitive</param>
        /// <remarks></remarks>
        public IniReader(string iniFilePath, bool isCaseSensitive, bool pIdentifySections) : this(iniFilePath, System.Text.Encoding.Default, DEFAULT_KeyValuePairDelimiter, DEFAULT_LineDelimiter, isCaseSensitive, pIdentifySections, new string[] { })
        {
        }

        /// <summary>
        /// Reads Ini File. Uses ANSII - WINDOWS  Encoding.  System.Text.Encoding.Default
        /// </summary>
        /// <param name="iniFilePath">The Ini File to read</param>
        /// <param name="KeyVDelimiter">The key value delimiter like Equals sign [ = ]</param>
        /// <param name="LineVDelimiter">The Line delimiter or Entry delimiter like New Line [ \n ]</param>
        /// <param name="isCaseSensitive">Indicate if fetching result will be case sensitive</param>
        /// <param name="IgnoreSpecialComments">Entries of Special Comments Starter lines like REM. Semi-colon ; </param>
        /// <remarks></remarks>
        public IniReader(string iniFilePath, string KeyVDelimiter, string LineVDelimiter, bool isCaseSensitive, bool pIdentifySections, params string[] IgnoreSpecialComments) : this(iniFilePath, System.Text.Encoding.Default, KeyVDelimiter, LineVDelimiter, isCaseSensitive, pIdentifySections, IgnoreSpecialComments)
        {
        }
        /// <summary>
        /// Reads Ini File
        /// </summary>
        /// <param name="iniFilePath">The Ini File to read</param>
        /// <param name="KeyVDelimiter">The key value delimiter like Equals sign [ = ]</param>
        /// <param name="LineVDelimiter">The Line delimiter or Entry delimiter like New Line [ \n ]</param>
        /// <param name="isCaseSensitive">Indicate if fetching result will be case sensitive</param>
        /// <param name="IgnoreSpecialComments">Entries of Special Comments Starter lines like REM. Semi-colon ; </param>
        /// <remarks></remarks>
        public IniReader(string iniFilePath, System.Text.Encoding encode, string KeyVDelimiter, string LineVDelimiter, bool isCaseSensitive, bool pIdentifySections, params string[] IgnoreSpecialComments)
        {
            try
            {
                string f = File.ReadAllText(iniFilePath, encode);
                if ((f.Trim() ?? "") != (string.Empty ?? ""))
                {
                    LoadData(f, KeyVDelimiter, LineVDelimiter, isCaseSensitive, pIdentifySections, new List<string>(IgnoreSpecialComments));
                }
            }
            catch (Exception e)
            {

                // REM All exceptions means it is a crazy file :D
                Modules.basMain.MyLogFile.Write(new EException("Error While Reading File: " + iniFilePath, e));
            }
        }


        /// <summary>
        /// Reads Ini File
        /// </summary>
        /// <param name="iniFileContents">Parse the contents of ini file to read</param>
        /// <param name="KeyVDelimiter">The key value delimiter like Equals sign [ = ]</param>
        /// <param name="LineVDelimiter">The Line delimiter or Entry delimiter like New Line [ \n ]</param>
        /// <param name="isCaseSensitive">Indicate if fetching result will be case sensitive</param>
        /// <param name="IgnoreSpecialComments">Entries of Special Comments Starter lines like REM. Semi-colon ; </param>
        /// <remarks></remarks>
        public IniReader(string iniFileContents, string KeyVDelimiter, string LineVDelimiter, List<string> IgnoreSpecialComments, bool isCaseSensitive, bool pIdentifySections)
        {
            LoadData(iniFileContents, KeyVDelimiter, LineVDelimiter, isCaseSensitive, pIdentifySections, IgnoreSpecialComments);
        }


        #endregion



        #region Properties

        private Dictionary<string, string> IniDetails;
        private bool ___isCaseSensitive;

        public bool isCaseSensitive()
        {
            return ___isCaseSensitive;
        }

        private bool _readSuccessFully = false;
        public const string DEFAULT_KeyValuePairDelimiter = "=";
        public const string DEFAULT_LineDelimiter = @"\r\n";
        private string KeyValuePairDelimiter = DEFAULT_KeyValuePairDelimiter;


        /// <summary>
        /// It is use to identify sections if Identify Sections was indicated
        /// </summary>
        /// <remarks></remarks>
        public const string SECTION__NAME___SEPERATOR = "__________SEC_________";
        private bool _______IsSectionIdentified;

        public bool IsSectionIdentified
        {
            get
            {
                return _______IsSectionIdentified;
            }
        }

        public bool isValid()
        {
            return _readSuccessFully;
        }

        public int KeysCounts()
        {
            if (!isValid())
                return 0;
            return Keys().Count();
        }

        public IEnumerable<string> Keys()
        {
            if (!isValid())
                return new List<string>();
            return IniDetails.Keys.AsEnumerable();
        }

        #endregion


        #region Methods


        public string getValue(string Key)
        {
            if (!___isCaseSensitive)
                Key = Key.ToLower();
            if (Keys().Contains(Key))
                return IniDetails[Key];
            return string.Empty;
        }

        private void LoadData(string iniFileContents, string KeyVDelimiter, string LineVDelimiter, bool isCaseSensitive, bool pIdentifySections, List<string> IgnoreSpecialComments)
        {
            KeyValuePairDelimiter = KeyVDelimiter;

            // REM Translate LineDelimiter
            if (Objects.EStrings.isEscapeCharacters(LineVDelimiter))
                LineVDelimiter = Objects.EStrings.translateEscapeCharacters(LineVDelimiter);
            if (IgnoreSpecialComments is null)
                IgnoreSpecialComments = new List<string>();
            var IgnoreComments = new List<string>(50);  // REM Maximum 50 types is enough :P
            IgnoreComments.Add(";");
            if (IgnoreSpecialComments.Contains(IgnoreComments[0]))
                IgnoreComments.RemoveAt(0); // REM avoid duplicates
            IgnoreComments.AddRange(IgnoreSpecialComments);
            IniDetails = new Dictionary<string, string>(100);  // REM 100 keys maximum lol :D
            ___isCaseSensitive = isCaseSensitive;
            _______IsSectionIdentified = pIdentifySections;
            try
            {
                string f = iniFileContents;
                if ((f.Trim() ?? "") != (string.Empty ?? ""))
                {

                    // REM Parse File
                    var Lines = f.Split(new string[] { LineVDelimiter }, StringSplitOptions.RemoveEmptyEntries);
                    if (Lines.Length > 0)
                    {
                        string queries = string.Empty;
                        // REM remove comments first
                        foreach (string Line in Lines)
                        {
                            bool IgnoreLine = false;
                            // REM support comments and special comments
                            foreach (string comment in IgnoreComments)
                            {
                                if (Line.StartsWith(comment, StringComparison.CurrentCultureIgnoreCase))
                                {
                                    IgnoreLine = true;
                                    break;
                                }
                            }

                            if (!IgnoreLine)
                                queries += Line + LineVDelimiter; // REM preserve the breaks
                        }




                        // If KeySections are meant to be identified
                        string vLastKeyIdentified = string.Empty;


                        // REM Now process the files using the real delimiter
                        foreach (string Line in queries.Split(new string[] { LineVDelimiter }, StringSplitOptions.RemoveEmptyEntries))
                        {
                            string pTestKey = Line.Trim(new char[] { ' ', '\t', '\r' });
                            if (pIdentifySections && pTestKey.StartsWith("[") && pTestKey.EndsWith("]"))
                            {
                                // This entry is a key
                                vLastKeyIdentified = pTestKey.Substring(1, pTestKey.Length - 2);
                                continue;
                            }

                            string vAppendSectionToKeyName = string.Empty;
                            if ((vLastKeyIdentified ?? "") != (string.Empty ?? ""))
                                vAppendSectionToKeyName = vLastKeyIdentified + SECTION__NAME___SEPERATOR;



                            // REM get keys values
                            // REM Dim keyValue As String() = Line.Split(New String() {KeyValuePairDelimiter}, StringSplitOptions.RemoveEmptyEntries)
                            var keyValue = Strings.Split(Line, KeyValuePairDelimiter, 2);
                            if (keyValue.Length == 2)
                            {

                                // REM if error occured here, it means user entered same key more than once
                                if (!___isCaseSensitive)
                                {
                                    IniDetails.Add(vAppendSectionToKeyName + keyValue[0].ToLower().Trim(new char[] { ' ', '\t', '\r' }), keyValue[1].Trim());
                                }
                                else
                                {
                                    IniDetails.Add(vAppendSectionToKeyName + keyValue[0].Trim(new char[] { ' ', '\t', '\r' }), keyValue[1].Trim());
                                }
                            }
                        }




                        // REM Reconfirm additions
                        if (IniDetails.Count != 0)
                            _readSuccessFully = true;
                    }
                }
            }
            catch (Exception e)
            {

                // REM All exceptions means it is a crazy file :D
                Modules.basMain.MyLogFile.Write(new EException(e));
            }
        }


        #endregion



    }
}