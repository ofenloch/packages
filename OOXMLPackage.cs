using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.IO.Packaging; // dotnet add package System.IO.Packaging
using System.Diagnostics;
using System.Xml;
using System.Xml.Linq;

namespace packages
{
    public class OOXMLPackage
    {
        /**
         * Core document relationship type.
         */
        public static string CORE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        /**
         * Custom properties relationship type.
         */
        public static string CUSTOM_PROPERTIES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
        /**
         * Visio 2010 VSDX equivalent of package {@link #CORE_DOCUMENT}
         */
        public static string VISIO_CORE_DOCUMENT = "http://schemas.microsoft.com/visio/2010/relationships/document";
        public static string VISIO_PAGES = "http://schemas.microsoft.com/visio/2010/relationships/pages";
        public static string VISIO_PAGE = "http://schemas.microsoft.com/visio/2010/relationships/page";

        string _packageFileName = "";
        Package _package = null;
        FileMode _fileMode = FileMode.OpenOrCreate;
        FileAccess _fileAccess = FileAccess.ReadWrite;
        OOXMLPackage(string fileName)
        {
            _packageFileName = fileName;
            try
            {
                _package = Package.Open(_packageFileName, FileMode.Open, FileAccess.ReadWrite);
            }
            catch (Exception e)
            {
                Console.WriteLine("Caught exception opening package in file \"{0}\": ", _packageFileName);
                Console.WriteLine(e.ToString());

            }
        } // OOXMLPackage(string fileName)
    } // public class OOXMLPackage
} // namespace packages