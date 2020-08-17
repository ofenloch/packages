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

        // loop over all package parts and show som information
        public static void IteratePackageParts(Package package)
        {
            // The way to get a reference to a package part is
            // by using its URI. Thus, we're reading the URI
            // for each part in the package.
            PackagePartCollection packageParts = package.GetParts();
            foreach (PackagePart packagePart in packageParts)
            {
                Console.WriteLine("Package part: {0}", packagePart.Uri);
                Console.WriteLine("  Content Type: {0}", packagePart.ContentType.ToString());
                if (packagePart.ContentType != "application/vnd.openxmlformats-package.relationships+xml")
                {
                    PackageRelationshipCollection packageRelationships = packagePart.GetRelationships();
                    foreach (PackageRelationship packageRelationship in packageRelationships)
                    {
                        Console.WriteLine("  PackageRelationship: id {0}", packageRelationship.Id);
                        Console.WriteLine("   Source URI  : {0}", packageRelationship.SourceUri);
                        Console.WriteLine("   Target URI  : {0}", packageRelationship.TargetUri);
                        Console.WriteLine("   Target Mode : {0}", packageRelationship.TargetMode);
                        Console.WriteLine("   Relationship Type: {0}", packageRelationship.RelationshipType.ToString());
                    }
                }
                else
                {
                    Console.WriteLine("  PackageRelationship parts cannot have relationships to other parts.");
                }
            }
        } // public static void IteratePackageParts(Package package)

        // unpack the given Pakage and write its content as XML files to the given target directory
        public static void UnpackPackage(Package package, string targetDir)
        {
            CreateDirectory(targetDir);
            // The way to get a reference to a package part is
            // by using its URI. Thus, we're reading the URI
            // for each part in the package.
            PackagePartCollection packageParts = package.GetParts();
            foreach (PackagePart packagePart in packageParts)
            {
                Uri uri = packagePart.Uri;
                Console.WriteLine("Package part: {0}", uri);
                string fileName = targetDir + uri;
                string dirName = Path.GetDirectoryName(fileName);
                CreateDirectory(dirName);
                Console.WriteLine("  file {0}", fileName);
                if (packagePart.ContentType.EndsWith("xml"))
                {
                    // Open the XML from the Page Contents part.
                    XDocument packagePartXML = GetXMLFromPart(packagePart);
                    packagePartXML.Save(fileName);
                }
                else
                {
                    // just save the non XML as it is
                    FileStream newFileStrem = new FileStream(fileName, FileMode.Create);
                    packagePart.GetStream().CopyTo(newFileStrem);
                }
            }
        } // public static void UnpackPackage(Package package, string targetDir)

        // get the XML document in a package part
        public static XDocument GetXMLFromPart(PackagePart packagePart)
        {
            XDocument partXml = null;
            // Open the packagePart as a stream and then 
            // open the stream in an XDocument object.
            Stream partStream = packagePart.GetStream();
            partXml = XDocument.Load(partStream);
            // close the part's stream so we won't get an exception when writing to it
            partStream.Close();
            return partXml;
        } // public static XDocument GetXMLFromPart(PackagePart packagePart)

        // create the given directory / path
        public static void CreateDirectory(string path)
        {
            try
            {
                // Determine whether the directory exists.
                if (Directory.Exists(path))
                {
                    // TODO: use logger Console.WriteLine("That path \"{0}\" exists already.", path);
                    return;
                }

                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                // TODO: use logger Console.WriteLine("The directory \"{0}\" was created successfully at {1}.", di.FullName, Directory.GetCreationTime(path));
            }
            catch (Exception e)
            {
                Console.WriteLine("CreateDirectory failed: {0}", e.ToString());
            }
            finally { }
        } // public static void CreateDirectory(string path)

        // get a specific PackagePart by its PackageRelationship (without having to iterate over all of the PackageParts)
        public static PackagePart GetPackagePartByRelationship(Package filePackage, string relationship)
        {

            // Use the namespace that describes the relationship 
            // to get the relationship.
            PackageRelationship packageRel =
                filePackage.GetRelationshipsByType(relationship).FirstOrDefault();
            PackagePart part = null;
            // If the Visio file package contains this type of relationship with 
            // one of its parts, return that part.
            if (packageRel != null)
            {
                // Clean up the URI using a helper class and then get the part.
                Uri docUri = PackUriHelper.ResolvePartUri(
                    new Uri("/", UriKind.Relative), packageRel.TargetUri);
                part = filePackage.GetPart(docUri);
            }
            return part;
        } // public static PackagePart GetPackagePartByRelationship(Package filePackage, string relationship)

        // get a specific PackagePart by its PackageRelationship to another PackagePart
        public static PackagePart GetPackagePartByRelationship(Package filePackage, PackagePart sourcePart, string relationship)
        {
            // This gets only the first PackagePart that shares the relationship
            // with the PackagePart passed in as an argument. You can modify the code
            // here to return a different PackageRelationship from the collection.
            PackageRelationship packageRel =
                sourcePart.GetRelationshipsByType(relationship).FirstOrDefault();
            PackagePart relatedPart = null;
            if (packageRel != null)
            {
                // Use the PackUriHelper class to determine the URI of PackagePart
                // that has the specified relationship to the PackagePart passed in
                // as an argument.
                Uri partUri = PackUriHelper.ResolvePartUri(
                    sourcePart.Uri, packageRel.TargetUri);
                relatedPart = filePackage.GetPart(partUri);
            }
            return relatedPart;
        } // public static PackagePart GetPackagePartByRelationship(Package filePackage, PackagePart sourcePart, string relationship)



        public static IEnumerable<XElement> GetXElementsByName(XDocument packagePart, string elementType)
        {
            // Construct a LINQ query that selects elements by their element type.
            IEnumerable<XElement> elements =
                from element in packagePart.Descendants()
                where element.Name.LocalName == elementType
                select element;
            // Return the selected elements to the calling code.
            return elements.DefaultIfEmpty(null);
        } // public static IEnumerable<XElement> GetXElementsByName(XDocument packagePart, string elementType)

        public static XElement GetXElementByAttribute(IEnumerable<XElement> elements, string attributeName, string attributeValue)
        {
            // Construct a LINQ query that selects elements from a group
            // of elements by the value of a specific attribute.
            IEnumerable<XElement> selectedElements =
                from el in elements
                where el.Attribute(attributeName).Value == attributeValue
                select el;
            // If there aren't any elements of the specified type
            // with the specified attribute value in the document,
            // return null to the calling code.
            try
            {
                // this throws an exception if the query doesn't return any resulte
                return selectedElements.DefaultIfEmpty(null).FirstOrDefault();
            }
            catch (Exception e)
            {
                // so we return simply null
                return null;
            }
        } // public static XElement GetXElementByAttribute(IEnumerable<XElement> elements, string attributeName, string attributeValue)

        //  --------------------------- CopyStream ---------------------------
        /// <summary>
        ///   Copies data from a source stream to a target stream.</summary>
        /// <param name="source">
        ///   The source stream to copy from.</param>
        /// <param name="target">
        ///   The destination stream to copy to.</param>
        public static void CopyStream(Stream source, Stream target)
        {
            const int bufSize = 0x1000;
            byte[] buf = new byte[bufSize];
            int bytesRead = 0;
            while ((bytesRead = source.Read(buf, 0, bufSize)) > 0)
                target.Write(buf, 0, bytesRead);
        }// end:CopyStream()

        // save the changed XML back to the package
        public static void SaveXDocumentToPart(PackagePart packagePart, XDocument partXML)

        {
            /**
                The following code uses the XmlWriter class and XmlWriterSettings class to write the XML back to 
                the package part. Although you can use the Save() method to save the XML back to the part, the 
                XmlWriter and XmlWriterSettings classes allow you finer control over the output, including specifying 
                the type of encoding. The XDocument class exposes a WriteTo(XmlWriter) method that takes an XmlWriter 
                object and writes XML back to a stream.
            **/
            // Create a new XmlWriterSettings object to 
            // define the characteristics for the XmlWriter
            XmlWriterSettings partWriterSettings = new XmlWriterSettings();
            partWriterSettings.Encoding = Encoding.UTF8;
            // Create a new XmlWriter and then write the XML back to the document part.
            XmlWriter partWriter = XmlWriter.Create(packagePart.GetStream(), partWriterSettings);
            // The GetStream() above throws an exception if the package part's stream is still open:
            // System.IO.IOException: Entries cannot be opened multiple times in Update mode.
            //   at System.IO.Compression.ZipArchiveEntry.OpenInUpdateMode()
            //   at System.IO.Compression.ZipArchiveEntry.Open()
            //   at System.IO.Packaging.ZipStreamManager.Open(ZipArchiveEntry zipArchiveEntry, FileMode streamFileMode, FileAccess streamFileAccess)
            //   at System.IO.Packaging.ZipPackagePart.GetStreamCore(FileMode streamFileMode, FileAccess streamFileAccess)
            //   at System.IO.Packaging.PackagePart.GetStream(FileMode mode, FileAccess access)
            //   at System.IO.Packaging.PackagePart.GetStream()
            //   at VisioEditor.Program.SaveXDocumentToPart(PackagePart packagePart, XDocument partXML) in C: \Users\Ofenloch.ol\VisualStudio\VisioEditor\VisioEditor\Program.cs:line 216
            //   at VisioEditor.Program.Main(String[] args) in C: \Users\Ofenloch.ol\VisualStudio\VisioEditor\VisioEditor\Program.cs:line 60
            partXML.WriteTo(partWriter);
            // Flush and close the XmlWriter.
            partWriter.Flush();
            partWriter.Close();
        } // public static void SaveXDocumentToPart(PackagePart packagePart, XDocument partXML)

        // make Visio recalculate the entire document when it is opened
        // by setting the RecalcDocument property (if it is not already set)
        public static void RecalcDocument(Package filePackage)
        {
            // Get the Custom File Properties part from the package and
            // and then extract the XML from it.
            PackagePart customPart = GetPackagePartByRelationship(filePackage, OOXMLPackage.CUSTOM_PROPERTIES);
            XDocument customPartXML = OOXMLPackage.GetXMLFromPart(customPart);
            // Check to see whether document recalculation has already been 
            // set for this document. If it hasn't, use the integer
            // value returned by CheckForRecalc as the property ID.
            int propertyID = CheckForRecalc(customPartXML);
            if (propertyID > -1)
            {
                XElement customPartRoot = customPartXML.Elements().ElementAt(0);
                // Two XML namespaces are needed to add XML data to this 
                // document. Here, we're using the GetNamespaceOfPrefix and 
                // GetDefaultNamespace methods to get the namespaces that 
                // we need. You can specify the exact strings for the 
                // namespaces, but that is not recommended.
                XNamespace customVTypesNS = customPartRoot.GetNamespaceOfPrefix("vt");
                XNamespace customPropsSchemaNS = customPartRoot.GetDefaultNamespace();
                // Construct the XML for the new property in the XDocument.Add method.
                // This ensures that the XNamespace objects will resolve properly, 
                // apply the correct prefix, and will not default to an empty namespace.
                customPartRoot.Add(
                    new XElement(customPropsSchemaNS + "property",
                        new XAttribute("pid", propertyID.ToString()),
                        new XAttribute("name", "RecalcDocument"),
                        new XAttribute("fmtid",
                            "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
                        new XElement(customVTypesNS + "bool", "true")
                    ));
            }
            // Save the Custom Properties package part back to the package.
            SaveXDocumentToPart(customPart, customPartXML);
        } // public static void RecalcDocument(Package filePackage)



        // check if the property RecalcDocument is already set
        // return its id if so or -1 if it is not set
        public static int CheckForRecalc(XDocument customPropsXDoc)
        {
            // Set the inital propertyID to -1, which is not an allowed value.
            // The calling code tests to see whether the pidValue is 
            // greater than -1.
            int propertyID = -1;
            // Get all of the property elements from the document. 
            IEnumerable<XElement> props = GetXElementsByName(
                customPropsXDoc, "property");
            // Get the RecalcDocument property from the document if it exists already.
            XElement recalcProp = GetXElementByAttribute(props,
                "name", "RecalcDocument");
            // If there is already a RecalcDocument instruction in the 
            // Custom File Properties part, then we don't need to add another one. 
            // Otherwise, we need to create a unique pid value.
            if (recalcProp != null)
            {
                return propertyID;
            }
            else
            {
                // Get all of the pid values of the property elements and then
                // convert the IEnumerable object into an array.
                IEnumerable<string> propIDs =
                    from prop in props
                    where prop.Name.LocalName == "property"
                    select prop.Attribute("pid").Value;
                string[] propIDArray = propIDs.ToArray();
                // Increment this id value until a unique value is found.
                // This starts at 2, because 0 and 1 are not valid pid values.
                int id = 2;
                while (propertyID == -1)
                {
                    if (propIDArray.Contains(id.ToString()))
                    {
                        id++;
                    }
                    else
                    {
                        propertyID = id;
                    }
                }
            }
            return propertyID;
        } // public static int CheckForRecalc(XDocument customPropsXDoc)

    } // public class OOXMLPackage
} // namespace packages