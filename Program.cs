using System;
using System.IO;
using System.IO.Packaging;

namespace packages
{
    class Program
    {
        static void Main(string[] args)
        {
            //string fileName = "./data/60489.vsdx";
            string fileName = "data/hello-world-unsigned.docx";
            try
            {
                Console.WriteLine("Opening the Package in file \"{0}\" ...", fileName);
                // We're not going to do any more than open
                // and read the list of parts in the package, although
                // we can create a package or read/write what's inside.
                using (Package fPackage = Package.Open(
                    fileName, FileMode.Open, FileAccess.Read))
                {

                    // The way to get a reference to a package part is
                    // by using its URI. Thus, we're reading the URI
                    // for each part in the package.
                    PackagePartCollection fParts = fPackage.GetParts();
                    foreach (PackagePart part in fParts)
                    {
                        Console.WriteLine("Package part: {0}", part.Uri);
                        Console.WriteLine("  Content Type: {0}", part.ContentType.ToString());
                        if (part.ContentType != "application/vnd.openxmlformats-package.relationships+xml")
                        {
                            PackageRelationshipCollection relationships = part.GetRelationships();
                            foreach (PackageRelationship relationship in relationships)
                            {
                                Console.WriteLine("  Relationship Type: {0}", relationship.RelationshipType.ToString());
                            }
                        }
                        else
                        {
                            Console.WriteLine("  PackageRelationship parts cannot have relationships to other parts.");
                        }
                    }
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
    } // class Program
} // namespace packages
