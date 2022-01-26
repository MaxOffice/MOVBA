
function Get-MVPackage {
    <#
.SYNOPSIS
    List the packages in the current library.
.DESCRIPTION
    The Get-MVPackage cmdlet lists the packages in the current 
    library. It only works if invoked from a library root directory.

    Without parameters, it lists all packages. Otherwise, it lists
    packages matching the names provided. Matching is case-insensitive.
    Invalid package names passed in this parameter are ignored.
.EXAMPLE
    Get-MVPackage

    This lists all packages.
.EXAMPLE
    Get-MVPackage MyAwesomePackage

    This lists a package called MyAwesomePackage, if it is available in
    the current library.
#>
    [CmdletBinding()]
    [OutputType([MVPackage[]])]
    Param (
        # Specifies names of packages that this cmdlet lists.
        # Wildcard patterns are not allowed. Names are 
        # case-insensitive.
        [Parameter(
            Mandatory = $false,
            Position = 0,
            ValueFromRemainingArguments = $true
        )]
        [string[]]
        $Name
    )

    try {
        [MVLibrary]::TestCurrentDirectory()

        If ($Name.Count -eq 1) {
            return [MVLibrary]::GetPackage($Name[0])
        }

        return [MVLibrary]::GetPackages($Name)
    }
    catch {
        $PSCmdlet.WriteError([MVLibrary]::CmdletError($_))
    }
}

function  New-MVPackage {
    <#
.SYNOPSIS
    Create a new package in the current library.
.DESCRIPTION
    The New-MVPackage cmdlet creates a new package in the current
    library. It only works if invoked from a library root directory.

    It checks if a package with the supplied name already exists. If
    not, it creates a directory with that name under the packages 
    directory of the library. If the -HasCustomUI parameter is 
    provided, it creates a directory called CustomUI under the 
    package directory, and creates an Office Open XML custom ui
    manifest file in it called customUI.xml.
.EXAMPLE
    New-MVPackage NewPackage1

    This creates a new package without a custom ui subdirectory.
.EXAMPLE
    New-MVPackage NewPackage1 -HasCustomUI

    This creates a new package with a custom ui subdirectory.
#>

    [CmdletBinding()]
    [OutputType('MVPackage')]
    param (
        # Specifies the name of the new package.
        [Parameter(Mandatory = $true, Position = 0)]
        [string]
        $Name,
        # Causes the CustomUI directory and file
        # to be created.
        [Parameter(Position = 1)]
        [switch]
        $HasCustomUI
    )
    
    try {
        [MVLibrary]::TestCurrentDirectory()

        if ("" -eq $Name) {
            Write-Error "Package name must be specifed." -ErrorAction Stop
        }

        $newPackage = [MVLibrary]::NewPackage($Name, $HasCustomUI)

        return $newPackage
    }
    catch {
        $PSCmdlet.WriteError([MVLibrary]::CmdletError($_))
    }
}

function Remove-MVPackage {
    <#
.SYNOPSIS
    Remove a package from the current library.
.DESCRIPTION
    The Remove-MVPackage cmdlet removes a package from the current
    library. It only works if invoked from a library root directory.
.EXAMPLE
    Remove-MVPackage OldPackage1

    This removes a package called OldPackage1.
#>
    [CmdletBinding()]
    param (
        # This specifies the name of the package to remove. The name
        # is case-insensitive.
        [Parameter(Position = 0)]
        [string]
        $PackageName = ""
    )

    try {
        [MVLibrary]::TestCurrentDirectory()

        if ("" -eq $PackageName) {
            Write-Error "Package name must be specifed." -ErrorAction Stop
        }

        [void] [MVLibrary]::RemovePackage($PackageName)
    }
    catch {
        $PSCmdlet.WriteError([MVLibrary]::CmdletError($_))
    }
}

function Get-MVLibrary {
    <#
.SYNOPSIS
    Get details of the current library.
.DESCRIPTION
    The Get-MVLibrary cmdlet gets the properties of the current library.
    It only works if invoked from a library root directory.
.EXAMPLE
    Get-MVLibrary

    This shows details of the current library.
#>
    [CmdletBinding()]
    param()

    try {
        [MVLibrary]::GetProjectConfig()
    }
    catch {
        $PSCmdlet.WriteError([MVLibrary]::CmdletError($_))
    }
}

function New-MVLibrary {
    <#
.SYNOPSIS
    Create a new VBA macro library.
.DESCRIPTION
    The New-MVLibrary cmdlet creates a new Microsoft Office 
    VBA macro library, which will contain VBA macros created 
    for a specific Microsoft Office application. Macros have 
    to be grouped into packages. Currently, a library can 
    contain macros for any one of: Microsoft Word, Excel and
    PowerPoint.
    
    The cmdlet creates the library directory, with two
    subdirectories called packages and out. It also creates 
    a file called library.json, and a .gitignore file that 
    ignores Microsoft Office temporary files and the out 
    subdirectory.
.EXAMPLE
    New-MVLibrary AwesomeLibrary

    This creates a Library directory named AwesomeLibrary
    under the current directory.
#>
    [CmdletBinding()]
    [OutputType('MVLibrary')]
    param (
        # This specifies the name of the new library.
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,
        # This specifies the Microsoft Office application that 
        # hosts the VBA macros in this library. Currently
        # supported values are Word, Excel and  PowerPoint.
        # Case insensitive.
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateScript({
                [MVLibrary]::TestOfficeApplication($_)
            })]
        [string]
        $Application
    )
    
    try {
        [MVLibrary]::Init($Name, $Application)
    }
    catch {
        $PSCmdlet.WriteError([MVLibrary]::CmdletError($_))
    }
}

function Build-MVLibrary {
    <#
.SYNOPSIS
    Build a Microsoft Office document or add-in from the current library.
.DESCRIPTION
    The Build-MVLibrary cmdlet builds a Microsoft Office document or 
    add-in from the current library. It only works if invoked from a
    library root directory.

    The actual file built will depend on the Application property
    of the library. In any case, the file will contain the combined
    macros and custom UI elements from the selected packages.

    By default, macros and custom UI elements of all packages in the
    library are selected. The -PackageNames parameter allows one or
    more packages to be selected specifically if all are not needed.
    Invalid package names passed in this parameter are ignored.

    The final file will be placed in the out subdirectory of the 
    library.
.EXAMPLE
    Build-MVLibrary AllMacros AddIn

    This builds an application-specific add-in file that contains
    all packages in the library.
.EXAMPLE
    Build-MVLibrary AllMacros Document Lucky1,Lucky2

    This builds an application-specific document file that contains
    the packages Lucky1 and Lucky2.
#>
    [CmdletBinding()]
    [OutputType('MVLibrary')]
    param (
        # This specifies the name of the document or add-in
        # file that will be built from packages in this
        # library. It should be a file name without an 
        # extension.
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $OutFileName,
        [Parameter(Mandatory = $false, Position = 1)]
        # This specifies the kind of output file to build.
        # Valid choices are AddIn or Document.
        [ValidateSet('AddIn', 'Document')]
        [string]
        $BuildType = 'AddIn',
        # This specifies the packages that will be built
        # into the output file. Not specifying this will
        # cause all the packages to be built it. Package 
        # names are case-insensitive.
        [Parameter(
            Mandatory = $false, 
            Position = 2,
            ValueFromRemainingArguments = $true
        )]
        [string[]]
        $PackageNames
    )
    
    try {
        [MVLibrary]::TestCurrentDirectory()

        $buildPackages = [MVLibrary]::GetPackages($PackageNames)

        if ($buildPackages.Count -eq 0) {
            Write-Error "No valid packages selected for build." -ErrorAction Stop
        }

        [MVLibrary]::Build($OutFileName, $BuildType, $buildPackages)
    }
    catch {
        $PSCmdlet.WriteError([MVLibrary]::CmdletError($_))
    }
}