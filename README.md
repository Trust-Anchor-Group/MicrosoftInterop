Microsoft Interoperation
===========================

This repository contains services that help with interoperation with Microsoft products,
using **Open XML SDK** libraries. This SDK allows for interoperation with documents saved using
the Open XML specification, not old legacy formats (i.e. `.docx` works, `.doc` not so).

## Projects

The solution contains the following C# projects:

| Project                        | Framework         | Description |
|:-------------------------------|:------------------|:------------|
| `TAG.Content.Microsoft`        | .NET Standard 2.0 | Class library for conversion of Microsoft artefacts (such as Word documents), using [Open XML SDK](https://sv.wikipedia.org/wiki/Office_Open_XML). |
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
| [Waher.Events](https://www.nuget.org/packages/Waher.Events/)                                       | An extensible architecture for event logging in the application. |
| [Waher.IoTGateway](https://www.nuget.org/packages/Waher.IoTGateway/)                               | Contains the [IoT Gateway](https://github.com/PeterWaher/IoTGateway) hosting environment. |
| [Waher.Networking](https://www.nuget.org/packages/Waher.Networking/)                               | Tools for working with communication, including troubleshooting. |
| [Waher.Networking.HTTP](https://www.nuget.org/packages/Waher.Networking.HTTP/)                     | Library for publishing information and services via HTTP. |
| [Waher.Runtime.Inventory](https://www.nuget.org/packages/Waher.Runtime.Inventory/)                 | Maintains an inventory of type definitions in the runtime environment, and permits easy instantiation of suitable classes, and inversion of control (IoC). |
| [Waher.Runtime.Text](https://www.nuget.org/packages/Waher.Runtime.Text/)                           | Tools for processing text, such as harmonized text maps, and comparing text differences. |

The Unit Tests further use the following libraries:

| Nuget                                                                                            | Description |
|:-------------------------------------------------------------------------------------------------|:------------|
| [Waher.Content.XML](https://www.nuget.org/packages/Waher.Content.XML/)                           | Library with tools for XML processing. |
| [Waher.Events.Console](https://www.nuget.org/packages/Waher.Events.Console/)                     | Outputs events logged to the console output. |
| [Waher.Runtime.Inventory.Loader](https://www.nuget.org/packages/Waher.Runtime.Inventory.Loader/) | Permits the inventory and seamless integration of classes defined in all available assemblies. |
