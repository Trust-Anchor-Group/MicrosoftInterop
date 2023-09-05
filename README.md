Microsoft Interoperation
===========================

This repository contains services that help with interoperation with Microsoft products,
using **Open XML SDK** libraries. This SDK allows for interoperation with documents saved using
the Open XML specification, not old legacy formats (i.e. `.docx` and `.xlsx` works, `.doc` or
`.xls` not so).

## Neuron-extensions

This solution demonstrates how to create the following type of extensions for the TAG Neuron:

* Installable service module package containing .NET service code, as well as file-based content.
* .NET-based web services.
* Internet Content Decoder for Word and Excel documents.
* Internet Content Converter for Word to Markdown.
* Dynamic extension of file-based Internet Content (Markdown, Javacript, CSS, etc.).
	* The MarkdownLab is extended to allow for uploading and conversion of Word documents.
	* The script Prompts is extended to allow for uploading and conversion of Excel documents.

## Projects

The solution contains the following C# projects:

| Project                        | Framework         | Description |
|:-------------------------------|:------------------|:------------|
| `TAG.Content.Microsoft`        | .NET Standard 2.0 | Class library for conversion of Microsoft artefacts (such as Word and Excel documents), using [Open XML SDK](https://sv.wikipedia.org/wiki/Office_Open_XML). |
| `TAG.Content.Microsoft.Test`   | .NET 6.0          | Unit tests for the `TAG.Content.Microsoft` library. |
| `TAG.Service.MicrosoftInterop` | .NET Standard 2.0 | Service module for the [TAG Neuron](https://lab.tagroot.io/Documentation/Index.md), providing web services for conversion of Microsoft documents. |

## Nugets

The following nugets external are used. They faciliate common programming tasks, and
enables the libraries to be hosted on an [IoT Gateway](https://github.com/PeterWaher/IoTGateway).
This includes hosting the bridge on the [TAG Neuron](https://lab.tagroot.io/Documentation/Index.md).
They can also be used standalone.

| Nuget                                                                                              | Description |
|:---------------------------------------------------------------------------------------------------|:------------|
| [Waher.Content](https://www.nuget.org/packages/Waher.Content/)                                     | Pluggable architecture for accessing, encoding and decoding Internet Content. |
| [Waher.Content.Markdown](https://www.nuget.org/packages/Waher.Content.Markdown/)                   | Library for parsing Markdown, and rendering it into different formats. |
| [Waher.Content.Markdown.Web](https://www.nuget.org/packages/Waher.Content.MarkdownWeb/)            | Library for converting Markdown documents to HTML in a web environment. |
| [Waher.Events](https://www.nuget.org/packages/Waher.Events/)                                       | An extensible architecture for event logging in the application. |
| [Waher.IoTGateway](https://www.nuget.org/packages/Waher.IoTGateway/)                               | Contains the [IoT Gateway](https://github.com/PeterWaher/IoTGateway) hosting environment. |
| [Waher.Networking](https://www.nuget.org/packages/Waher.Networking/)                               | Tools for working with communication, including troubleshooting. |
| [Waher.Networking.HTTP](https://www.nuget.org/packages/Waher.Networking.HTTP/)                     | Library for publishing information and services via HTTP. |
| [Waher.Runtime.Inventory](https://www.nuget.org/packages/Waher.Runtime.Inventory/)                 | Maintains an inventory of type definitions in the runtime environment, and permits easy instantiation of suitable classes, and inversion of control (IoC). |
| [Waher.Runtime.Text](https://www.nuget.org/packages/Waher.Runtime.Text/)                           | Tools for processing text, such as harmonized text maps, and comparing text differences. |
| [Waher.Security](https://www.nuget.org/packages/Waher.Security/)                                   | Library with assorted security utilities. |
| [Waher.Security.Users](https://www.nuget.org/packages/Waher.Security.Users/)                       | Security architecture based on users, roles and privileges. |

The Unit Tests further use the following libraries:

| Nuget                                                                                            | Description |
|:-------------------------------------------------------------------------------------------------|:------------|
| [Waher.Content.XML](https://www.nuget.org/packages/Waher.Content.XML/)                           | Library with tools for XML processing. |
| [Waher.Events.Console](https://www.nuget.org/packages/Waher.Events.Console/)                     | Outputs events logged to the console output. |
| [Waher.Runtime.Inventory.Loader](https://www.nuget.org/packages/Waher.Runtime.Inventory.Loader/) | Permits the inventory and seamless integration of classes defined in all available assemblies. |
| [Waher.Script](https://www.nuget.org/packages/Waher.Script/)                                     | Library with an extensible script processing architecture. |
| [Waher.Script.Content](https://www.nuget.org/packages/Waher.Script.Content/)                     | Extends the scripting environment with content-related script extensions. |

## Installable Package

The `TAG.Service.MicrosoftInterop` project has been made into a package that can be downloaded and installed on any 
[TAG Neuron](https://lab.tagroot.io/Documentation/Index.md).
To create a package, that can be distributed or installed, you begin by creating a *manifest file*. The
`TAG.Service.MicrosoftInterop` project has a manifest file called `TAG.Service.MicrosoftInterop.manifest`. It defines the
assemblies and content files included in the package. You then use the `Waher.Utility.Install` and `Waher.Utility.Sign` command-line
tools in the [IoT Gateway](https://github.com/PeterWaher/IoTGateway) repository, to create a package file and cryptographically
sign it for secure distribution across the Neuron network.

The Microsoft Interop service is published as a package on TAG Neurons. If your neuron is connected to this network, you can install the
package using the following information:

| Package information                                                                                                              ||
|:-----------------|:---------------------------------------------------------------------------------------------------------------|
| Package          | `TAG.MicrosoftInterop.package`                                                                                 |
| Installation key | `Y/0hf+O003/pMh6CDnQTowb3DMJj3X28Xu0H0/bOPsIdGo+XOGY2kWsEyxkpKMSNdAOjSGDlxUIA00c066163c7125123382bdd308a2ad35` |
| More Information | https://lab.tagroot.io/Community/Post/Microsoft_Interoperability_API                                           |

## Building, Compiling & Debugging

The repository assumes you have the [IoT Gateway](https://github.com/PeterWaher/IoTGateway) repository cloned in a folder called
`C:\My Projects\IoT Gateway`, and that this repository is placed in `C:\My Projects\MicrosoftInterop`. You can place the
repositories in different folders, but you need to update the build events accordingly. To run the application, you select the
`TAG.Service.MicrosoftInterop` project as your startup project. It will execute the console version of the
[IoT Gateway](https://github.com/PeterWaher/IoTGateway), and make sure the compiled files of the `MicrosoftInterop` solution
is run with it.

### Gateway.config

To simplify development, once the project is cloned, add a `FileFolder` reference
to your repository folder in your [gateway.config file](https://lab.tagroot.io/Documentation/IoTGateway/GatewayConfig.md). 
This allows you to test and run your changes to Markdown and Javascript immediately, 
without having to synchronize the folder contents with an external 
host, or recompile or go through the trouble of generating a distributable software 
package just for testing purposes. Changes you make in .NET can be applied in runtime
if you the *Hot Reload* permits, otherwise you need to recompile and re-run the
application again.

Example of how to point a web folder to your project folder:

```
<FileFolders>
  <FileFolder webFolder="/MicrosoftInterop" folderPath="C:\My Projects\MicrosoftInterop\TAG.Service.MicrosoftInterop\Root\MicrosoftInterop"/>
</FileFolders>
```

**Note**: Once file folder reference is added, you need to restart the IoT Gateway service for the change to take effect.

**Note 2**:  Once the gateway is restarted, the source for the files is in the new location. Any changes you make in the corresponding
`ProgramData` subfolder will have no effect on what you see via the browser.

**Note 3**: This file folder is only necessary on your developer machine, to give you real-time updates as you edit the files in your
developer folder. It is not necessary in a production environment, as the files are copied into the correct folders when the package 
is installed.
