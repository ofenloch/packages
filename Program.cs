using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace packages
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = "./data/60489.vsdx";
            //string fileName = "data/hello-world-unsigned.docx";
            try
            {
                Console.WriteLine("Opening the Package in file \"{0}\" ...", fileName);
                // We're not going to do any more than open
                // and read the list of parts in the package, although
                // we can create a package or read/write what's inside.
                using (Package fPackage = Package.Open(
                    fileName, FileMode.Open, FileAccess.Read))
                {
                    // we need this only to get info or for debugging / testing
                    //IteratePackageParts(fPackage);

                    // Get a reference to the Visio Document part contained in the file package.
                    PackagePart documentPart = GetPackagePart(fPackage,
                        "http://schemas.microsoft.com/visio/2010/relationships/document");
                    // Get a reference to the collection of pages in the document, 
                    // and then to the first page in the document.
                    PackagePart pagesPart = GetPackagePart(fPackage, documentPart,
                        "http://schemas.microsoft.com/visio/2010/relationships/pages");
                    PackagePart pagePart = GetPackagePart(fPackage, pagesPart,
                        "http://schemas.microsoft.com/visio/2010/relationships/page");
                    // Open the XML from the Page Contents part.
                    XDocument pageXML = GetXMLFromPart(pagePart);
                    // save the XML document representing the first page as XML file
                    pageXML.Save("./data/page1.xml");
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("Error: {0}", err.Message);
            }
            finally
            {
                Console.Write("\nPress any key to continue ...");
                Console.Write("\nBye!\n");
                Console.ReadKey();
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
            return partXml;
        } // private static XDocument GetXMLFromPart(PackagePart packagePart)


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

    } // class Program
} // namespace packages
