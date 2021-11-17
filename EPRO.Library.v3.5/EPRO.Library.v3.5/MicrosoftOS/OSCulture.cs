using System;
using System.Globalization;
using System.Text;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ELibrary.Standard.MicrosoftOS
{
    public class OSCulture
    {

        #region Properties


        private CultureInfo _____ClassCulture;

        public CultureInfo ClassCulture
        {
            get
            {
                return _____ClassCulture;
            }
        }

        public Cultures ClassCultureType
        {
            get
            {
                return (Cultures)Conversions.ToInteger(ClassCulture.LCID);
            }
        }

        #endregion


        #region Enums

        public enum Cultures
        {
            ENGLISH___________USA = 1033,
            ENGLISH___________UK = 2057,
            FRENCH_________FRANCE = 1036,
            UNKNOWN______________ = -1
        }

        #endregion


        #region Methods

        /// <summary>
        /// Throws Exception Exception for not supported
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public static OSCulture getOSCultureInstalledOnPC()
        {
            try
            {
                return new OSCulture((Cultures)Conversions.ToInteger(CultureInfo.InstalledUICulture.LCID));
            }
            catch (Exception ex)
            {
                throw new Exception("Please, check to see that this OS Culture is Supported: " + Environment.NewLine + getInstalledCultureSummary());
            }
        }


        /// <summary>
        /// Throws Exception Exception for not supported
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public static OSCulture getOSCultureForCurrentThread()
        {
            try
            {
                return new OSCulture((Cultures)Conversions.ToInteger(CultureInfo.CurrentUICulture.LCID));
            }
            catch (Exception ex)
            {
                throw new Exception("Please, check to see that this OS Culture is Supported: " + Environment.NewLine + getCurrentThreadCultureSummary());
            }
        }


        /// <summary>
        /// The summary of the OS Current UI Culture on this calling thread. It is usually same if user didn't create specific culture using  CultureInfo.CreateSpecificCulture("fr-FR")
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public string getClassCultureSummary()
        {
            return getCultureSummary(_____ClassCulture);
        }


        /// <summary>
        /// The summary of the OS Current UI Culture on this calling thread. It is usually same if user didn't create specific culture using  CultureInfo.CreateSpecificCulture("fr-FR")
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string getCurrentThreadCultureSummary()
        {
            return getCultureSummary(CultureInfo.CurrentUICulture);
        }


        /// <summary>
        /// The summary of the OS Installed Culture
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string getInstalledCultureSummary()
        {
            return getCultureSummary(CultureInfo.InstalledUICulture);
        }

        public static string getCultureSummary(CultureInfo pCulture)
        {
            var dCulture = pCulture;
            var result = new StringBuilder();
            result.AppendLine("Default Language Info: ");
            result.AppendLine("Name: " + Constants.vbTab + Constants.vbTab + Constants.vbTab + dCulture.Name);
            result.AppendLine("Display Name: " + Constants.vbTab + Constants.vbTab + Constants.vbTab + dCulture.DisplayName);
            result.AppendLine("English Name: " + Constants.vbTab + Constants.vbTab + Constants.vbTab + dCulture.EnglishName);
            result.AppendLine("2-letter ISO Name: " + Constants.vbTab + Constants.vbTab + dCulture.TwoLetterISOLanguageName);
            result.AppendLine("3-letter ISO Name: " + Constants.vbTab + Constants.vbTab + dCulture.ThreeLetterISOLanguageName);
            result.AppendLine("3-letter Win32 API Name: " + Constants.vbTab + dCulture.ThreeLetterWindowsLanguageName);
            result.AppendLine("LCID Identifier: " + Constants.vbTab + Constants.vbTab + dCulture.LCID);
            result.AppendLine("NativeName: " + Constants.vbTab + Constants.vbTab + Constants.vbTab + dCulture.NativeName);
            result.AppendLine("DateTimeFormat: " + Constants.vbTab + Constants.vbTab + dCulture.DateTimeFormat.ToString());
            return result.ToString();
        }

        #endregion




        #region Constructors

        /// <summary>
        /// Use ENGLISH___________USA 
        /// </summary>
        /// <remarks></remarks>
        public OSCulture() : this(Cultures.ENGLISH___________USA)
        {
        }


        /// <summary>
        /// If unknown is specified it switches to English
        /// </summary>
        /// <param name="pCulture"></param>
        /// <remarks></remarks>
        public OSCulture(Cultures pCulture)
        {
            if (pCulture != Cultures.UNKNOWN______________)
            {
                _____ClassCulture = new CultureInfo((int)pCulture);
            }
            else
            {
                _____ClassCulture = new CultureInfo((int)Cultures.ENGLISH___________USA);
            }
        }


        #endregion



    }
}