using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace ELibrary.Standard.AppConfigurations
{

    /// <summary>
    /// Purpose of this class is to read an XML value without error. Regardless of the arrangement of the elements.
    /// Please note that XML tag are case sensitive
    /// </summary>
    /// <remarks></remarks>
    public class SimpleXMLReader
    {
        public SimpleXMLReader(string pXMLContent)
        {
            try
            {
                ______ParsedXMLFile = XDocument.Parse(pXMLContent);
            }
            catch (Exception ex)
            {
                Program.Logger.Print(ex);
            }
        }

        private XDocument ______ParsedXMLFile;

        public XDocument ParsedXMLFile
        {
            get
            {
                return ______ParsedXMLFile;
            }
        }

        public bool IsValid
        {
            get
            {
                return ______ParsedXMLFile is object;
            }
        }





        /// <summary>
        /// Returns all occurrences of this TagElements
        /// </summary>
        /// <param name="pTagName"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public IEnumerable<XElement> getElementsByTagName(string pTagName)
        {
            if (!IsValid)
                return null;
            try
            {
                var p = ______ParsedXMLFile.Descendants(pTagName).ToList();
                return p;
            }
            catch (Exception ex)
            {
                Program.Logger.Print(ex);
                return null;
            }
        }


        /// <summary>
        /// get first occurence of this element
        /// </summary>
        /// <param name="pTagName"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public XElement getElementByTagName(string pTagName)
        {
            var p = getElementsByTagName(pTagName);
            if (p is object)
                return p.First();
            return null;
        }

        public IEnumerable<string> getElementsByTagNameValues(string pTagName)
        {
            var p = getElementsByTagName(pTagName);
            if (p is object)
                return p.Select(x => x.Value).ToList();
            return null;
        }

        public string getElementByTagNameValue(string pTagName)
        {
            var p = getElementByTagName(pTagName);
            if (p is object)
                return p.Value;
            return string.Empty;
        }
    }
}