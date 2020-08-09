# System.IO.Packaging Namespace

Start with a hello world console application in dotnet 2.2 on Kubuntu 14.04 LTS by running `dotnet new console --name packages`.

Add package **System.IO.Packaging** by running `dotnet add package System.IO.Packaging`.

```bash
ofenloch@kubuntu1404ofenloch:~/workspaces/dotnet/packages$ dotnet add package System.IO.Packaging
  Writing /tmp/tmpJ6XZpU.tmp
info : Adding PackageReference for package 'System.IO.Packaging' into project '/home/ofenloch/workspaces/dotnet/packages/packages.csproj'.
info : Restoring packages for /home/ofenloch/workspaces/dotnet/packages/packages.csproj...
info :   GET https://api.nuget.org/v3-flatcontainer/system.io.packaging/index.json
info :   OK https://api.nuget.org/v3-flatcontainer/system.io.packaging/index.json 471ms
info :   GET https://api.nuget.org/v3-flatcontainer/system.io.packaging/4.7.0/system.io.packaging.4.7.0.nupkg
info :   OK https://api.nuget.org/v3-flatcontainer/system.io.packaging/4.7.0/system.io.packaging.4.7.0.nupkg 32ms
info : Installing System.IO.Packaging 4.7.0.
info : Package 'System.IO.Packaging' is compatible with all the specified frameworks in project '/home/ofenloch/workspaces/dotnet/packages/packages.csproj'.
info : PackageReference for package 'System.IO.Packaging' version '4.7.0' added to file '/home/ofenloch/workspaces/dotnet/packages/packages.csproj'.
info : Committing restore...
info : Writing assets file to disk. Path: /home/ofenloch/workspaces/dotnet/packages/obj/project.assets.json
log  : Restore completed in 1.3 sec for /home/ofenloch/workspaces/dotnet/packages/packages.csproj.
ofenloch@kubuntu1404ofenloch:~/workspaces/dotnet/packages$
```

The goal is to manipulate Visio files. But the package stuff is common to all Office XML formats.

Informational stuff:

 * [Manipulate the Visio file format programmatically](https://docs.microsoft.com/en-us/office/client-developer/visio/how-to-manipulate-the-visio-file-format-programmatically)

 * [Introduction to the Visio file format (.vsdx)](https://docs.microsoft.com/en-us/office/client-developer/visio/introduction-to-the-visio-file-formatvsdx)

 * [VSDX: the new Visio file format](https://www.microsoft.com/en-us/microsoft-365/blog/2012/09/10/vsdx-the-new-visio-file-format/)