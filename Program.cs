using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.IO.Packaging; // dotnet add package System.IO.Packaging
using System.Diagnostics;
using System.Xml;
using System.Xml.Linq;

using packages;

namespace packages
{

    class Program
    {
        static string outputDirectory = "./testoutput/";
        static void Main(string[] args)
        {
            //string fileNameOrig = "./data/60489.vsdx";
            //string fileNameOrig = "data/hello-world-unsigned.docx";
            string fileNameOrig = "./data/Visio Package.vsdx";
            string fileNameCopy = outputDirectory + Path.GetFileName(fileNameOrig);
            try
            {
                CreateDirectory(outputDirectory);
                File.Copy(fileNameOrig, fileNameCopy, true);
                Console.WriteLine("Opening the Package in file \"{0}\" ...", fileNameCopy);

                // We're not going to do any more than open
                // and read the list of parts in the package, although
                // we can create a package or read/write what's inside.
                using (Package fPackage = Package.Open(fileNameCopy, FileMode.Open, FileAccess.ReadWrite))
                {
                    // we need this only to get info or for debugging / testing
                    //IteratePackageParts(fPackage);

                    // Get a reference to the Visio Document part contained in the file package.
                    PackagePart documentPart = GetPackagePart(fPackage, OOXMLPackage.VISIO_CORE_DOCUMENT);
                    if (documentPart != null)
                    {
                        // Get a reference to the collection of pages in the document, 
                        // and then to the first page in the document.
                        PackagePart pagesPart = GetPackagePart(fPackage, documentPart, OOXMLPackage.VISIO_PAGES);
                        if (pagesPart != null)
                        {
                            PackagePart page1Part = GetPackagePart(fPackage, pagesPart, OOXMLPackage.VISIO_PAGE);
                            if (page1Part != null)
                            {
                                // Open the XML from the Page Contents part.
                                XDocument page1XML = GetXMLFromPart(page1Part);
                                page1XML.Save(outputDirectory + "page1_orig.xml");
                                // Get all of the shapes from the page by getting
                                // all of the Shape elements from the pageXML document.
                                IEnumerable<XElement> shapesXML = GetXElementsByName(page1XML, "Shape");
                                if (shapesXML != null)
                                {
                                    // Select a Shape element from the shapes on the page by 
                                    // its name. You can modify this code to select elements
                                    // by other attributes and their values.
                                    XElement startEndShapeXML =
                                        GetXElementByAttribute(shapesXML, "NameU", "Start/End");
                                    if (startEndShapeXML != null)
                                    {
                                        Console.WriteLine("Found shape named \"Start/End\"");
                                        // Query the XML for the shape to get the Text element, and
                                        // return the first Text element node.
                                        IEnumerable<XElement> textElements = from element in startEndShapeXML.Elements()
                                                                             where element.Name.LocalName == "Text"
                                                                             select element;
                                        XElement textElement = textElements.ElementAt(0);
                                        // Change the shape text, leaving the <cp> element alone.
                                        textElement.LastNode.ReplaceWith("Start process\n");
                                        /* CAUTION
                                            In the previous code example, the existing shape text and the string used to replace it have the same 
                                            number of characters. Also note that the LINQ query changes the value of the last child node of the returned 
                                            element (which, in this case, is a text node). This is done to avoid changing the settings of the cp element 
                                            that is a child of the Text element. It is possible to cause file instability if you alter shape text 
                                            programmatically by overwriting all children of the Text element. As in the example above, the text 
                                            formatting is represented by cp elements under the Text element in the file. The definition of the formatting 
                                            is stored in the parent Section element. If these two pieces of information become inconsistent, then the file 
                                            may not behave as expected. Visio heals many inconsistencies, but it is better to ensure that any programmatic 
                                            changes are consistent so that you are controlling the ultimate behavior of the file.
                                        */

                                        // Insert a new Cell element in the Start/End shape that adds an arbitrary
                                        // local ThemeIndex value. This code assumes that the shape does not 
                                        // already have a local ThemeIndex cell.
                                        startEndShapeXML.Add(new XElement("Cell",
                                            new XAttribute("N", "ThemeIndex"),
                                            new XAttribute("V", "25"),
                                            new XProcessingInstruction("NewValue", "V")));
                                        // Change the shape's horizontal position on the page 
                                        // by getting a reference to the Cell element for the PinY 
                                        // ShapeSheet cell and changing the value of its V attribute.
                                        XElement pinYCellXML = GetXElementByAttribute(
                                            startEndShapeXML.Elements(), "N", "PinY");
                                        pinYCellXML.SetAttributeValue("V", "2");
                                        // Add instructions to Visio to recalculate the entire document
                                        // when it is next opened.
                                        RecalcDocument(fPackage);

                                        // Save the XML back to the Page Contents part.
                                        SaveXDocumentToPart(page1Part, page1XML);
                                        fPackage.Close();
                                        Console.WriteLine("Closed modified package in file \"{0}\"", fileNameCopy);
                                    }
                                    else
                                    {
                                        Console.WriteLine("Couldn't find shape named \"Start/End\"");
                                    }
                                }
                                // save the XML document representing the first page as XML file
                                CreateDirectory(outputDirectory);
                                page1XML.Save(outputDirectory + "page1_possibly_modified.xml");
                            } // if (page1Part != null )
                            else
                            {
                                Console.WriteLine("Couldn't find Visio page1Part.");
                            }
                        } // if (pagesPart != null)
                        else
                        {
                            Console.WriteLine("Couldn't find Visio pagesPart.");
                        }
                    } // if (documentPart != null)
                    else
                    {
                        Console.WriteLine("Couldn't find Visio documentPart. Trying to get Office document ...");
                        documentPart = GetPackagePart(fPackage, OOXMLPackage.CORE_DOCUMENT);
                        if (documentPart != null)
                        {
                            Console.WriteLine("Found Office documentPart with URI \"{0}\"", documentPart.Uri);
                        }
                        else
                        {
                            Console.WriteLine("Couldn't find Office documentPart.");
                        }
                    }
                } // using (Package fPackage = Package.Open(fileNameCopy, FileMode.Open, FileAccess.Read))

                // using (Package fPackage = Package.Open(fileNameCopy, FileMode.Open, FileAccess.Read))
                // {
                //     string targetDir = outputDirectory + Path.GetFileName(fileNameCopy) + ".unpacked";
                //     UnpackPackage(fPackage, targetDir);
                // }

            }
            catch (Exception err)
            {
                Console.WriteLine("Error: {0}", err.Message);
                Console.WriteLine(err.StackTrace);
            }
            finally
            {
                // Console.Write("\nPress any key to continue ...");
                // Console.Write("\nBye!\n");
                // Console.ReadKey();
            }
        } // static void Main(string[] args)

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
                        Console.WriteLine("  Relationship Type: {0}", packageRelationship.RelationshipType.ToString());
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

        // get a specific PackagePart by its PackageRelationship (without having to iterate over all of the PackageParts)
        private static PackagePart GetPackagePart(Package filePackage, string relationship)
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
        } // private static PackagePart GetPackagePart(Package filePackage, string relationship)

        // get a specific PackagePart by its PackageRelationship to another PackagePart
        private static PackagePart GetPackagePart(Package filePackage, PackagePart sourcePart, string relationship)
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
        } // private static PackagePart GetPackagePart(Package filePackage, PackagePart sourcePart, string relationship)

        // get the XML document in a package part
        private static XDocument GetXMLFromPart(PackagePart packagePart)
        {
            XDocument partXml = null;
            // Open the packagePart as a stream and then 
            // open the stream in an XDocument object.
            Stream partStream = packagePart.GetStream();
            partXml = XDocument.Load(partStream);
            // close the part's stream so we won't get an exception when writing to it
            partStream.Close();
            return partXml;
        } // private static XDocument GetXMLFromPart(PackagePart packagePart)

        private static IEnumerable<XElement> GetXElementsByName(XDocument packagePart, string elementType)
        {
            // Construct a LINQ query that selects elements by their element type.
            IEnumerable<XElement> elements =
                from element in packagePart.Descendants()
                where element.Name.LocalName == elementType
                select element;
            // Return the selected elements to the calling code.
            return elements.DefaultIfEmpty(null);
        } // private static IEnumerable<XElement> GetXElementsByName(XDocument packagePart, string elementType)

        private static XElement GetXElementByAttribute(IEnumerable<XElement> elements, string attributeName, string attributeValue)
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
        } // private static XElement GetXElementByAttribute(IEnumerable<XElement> elements, string attributeName, string attributeValue)

        //  --------------------------- CopyStream ---------------------------
        /// <summary>
        ///   Copies data from a source stream to a target stream.</summary>
        /// <param name="source">
        ///   The source stream to copy from.</param>
        /// <param name="target">
        ///   The destination stream to copy to.</param>
        private static void CopyStream(Stream source, Stream target)
        {
            const int bufSize = 0x1000;
            byte[] buf = new byte[bufSize];
            int bytesRead = 0;
            while ((bytesRead = source.Read(buf, 0, bufSize)) > 0)
                target.Write(buf, 0, bytesRead);
        }// end:CopyStream()

        // create the given directory / path
        private static void CreateDirectory(string path)
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
        } // private static void CreateDirectory(string path)

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
        private static void RecalcDocument(Package filePackage)
        {
            // Get the Custom File Properties part from the package and
            // and then extract the XML from it.
            PackagePart customPart = GetPackagePart(filePackage, CUSTOM_PROPERTIES);
            XDocument customPartXML = GetXMLFromPart(customPart);
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
        } // private static void RecalcDocument(Package filePackage)



        // check if the property RecalcDocument is already set
        // return its id if so or -1 if it is not set
        private static int CheckForRecalc(XDocument customPropsXDoc)
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
        } // private static int CheckForRecalc(XDocument customPropsXDoc)



    } // class Program
} // namespace packages
