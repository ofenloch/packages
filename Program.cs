using System;
using System.IO;
using System.IO.Packaging;

namespace packages
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = "./data/60489.vsdx";
            try
            {
                Console.WriteLine("Opening the VSDX file {0} ...", fileName);
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
                    foreach (PackagePart fPart in fParts)
                    {
                        Console.WriteLine("Package part: {0}", fPart.Uri);
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
