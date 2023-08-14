Microsoft Interoperation
===========================

This repository contains services that help with interoperation with Microsoft products. 
The services uses the `Microsoft.Interop.Office.*` nuget libraries to interoperate with the 
corresponding Office products. These need to be installed on the same machine, for the 
interoperation to work.

Security Notice
------------------

To avoid creating a vulnerability of injection by allowing processing of documents 
containing script, it is important that you **disable** Macro execution when documents are 
opened in Office on the machine. This has to be done manually after installing office, and 
before publishing any services that allow external users to process office documents using
the service. You can set the macro security level to **high** or **disabled** in the 
*Developer tab*, *Macro security level*.

Primary Interop Assemblies
------------------------------

During development, you must make sure the *Primary Interop Assemblies* are installed on
your computer. They are managed by *Visual Studio* itself. If you don't have these 
installed already, go to the Visual Studio installer, click *Modify*, and select the
*Office/SharePoint development* option, to make sure Office development tolls are installed
on the developer machine. The projects will assume the PIA assemblies have been installed
in the following folder:

```
C:\Program Files (x86)\Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15
```

If installed in another folder on your development machine, you might have to update the
references accordingly.


