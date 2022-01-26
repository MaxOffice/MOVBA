# MOVBA
A PowerShell module to manage VBA macro libraries.

![MOVBA Logo](/MOVBA.png)

## Concepts

In the context of this project, a VBA macro _library_ is a collection of useful
functionality implemented using VBA macros. A single library contains macros 
for one Microsoft Office application only. 

Within a library, macros are further arranged into _packages_. Each package implements a particular functionality. A package can expose its functionality via functions that can be called from other macros, or via an interactive user interface.

Packages in a library can be combined to form a Microsoft Office document or add-in. Users can choose all packages, or selective packages, when they build a document or add-in from the library.

## Libraries

A library is a directory with the following structure:

```
  \<library name>
  |
  --out\
  |
  --packages\
  |
  library.json
```

The `out\` subdirectory contains add-ins or documents created by combining packages in the library.

The `packages` subdirectory contains packages.

The `library.json` file contains metadata about the library, including the name of the library and the Microsoft Office application which is the target of all macros contained in the library.

The `Get-MVLibrary`, `New-MVLibrary` and `Build-MVLibrary` cmdlets in this module are used to manage libraries.

## Packages

A package is a directory with the following structure:

```
  \<package name>
  |
  --Tests\
  |
  --CustomUI\
  |   |
  |   customUI.xml
  |
  <package VBA files..>
```

A package must have a name that is unique in the library. Package names are treated case-insensitively in this module.

The package directory contains VBA files, such as VBA Modules (.bas files), VBA Class Modules (.cls files), and UserForms (.frm and .frx files). When a document or an add-in is built from the library, and a package is included, all these files are imported into the resulting document or add-in.

The `tests\` subdirectory should contain a document created using the Microsoft Office application that is the target of the parent library. This document should have test cases for checking the functionality implemented by the package.

The `CustomUI\` subdirectory contains a file called `customUI14.xml`, written using the Microsoft Office [RibbonX](https://docs.microsoft.com/en-us/office/vba/library-reference/concepts/customize-the-office-fluent-ribbon-by-using-an-open-xml-formats-file) specification. When a document or an add-in is built from the library, and a package is included, its `customUI14.xml` file is merged with that of other packages and imported into the resulting document or add-in.

The `Get-MVPackage`, `New-MVPackage` and `Remove-MVPackage` cmdlets in this module are used to manage packages.
