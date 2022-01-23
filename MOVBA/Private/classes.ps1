#region XML-processing Classes

# The following classes process different kinds of XML files found
# in an Office Open XML document.

Class xmlFile {
    [string] $ns = ""
    [hashtable] $nsm = @{}

    [xml] $xdoc
    [bool] hidden $isvalid = $false

    [void] CheckValidity() {
        If (-not $this.isValid) {
            Throw "The schema is not valid."
        }
    }

    [void] hidden setns([string]$rootNsUrn) {
        $this.ns = $rootNsUrn
        $this.nsm["x"] = $rootNsUrn         
    }

    [void] ValidateSchema() {
        $this.isvalid = $true
    }

    [void] Save([string] $filePath) {
        $this.CheckValidity()
        Set-Content -LiteralPath $filePath ($this.xdoc.InnerXml) -ErrorAction Stop
    }

    [void] SaveFormatted([string] $filePath) {
        $this.CheckValidity()
        $this.xdoc.Save($filePath)
    }
 
    xmlfile([string] $filePath, [string] $rootNsUrn) {
        $this.setns($rootNsUrn)
        $this.xdoc = [xml] (Get-Content $filePath -ErrorAction Stop)
        $this.ValidateSchema()
    }

    xmlfile([string] $rootNsUrn) {
        $this.setns($rootNsUrn)
        $this.xdoc = [xml]"<root xmlns='$rootNsUrn'></root>"
        $this.isvalid = $true
    }

    xmlfile() {
        $this.isvalid = $false
    }
}

Class contentTypesXMLFile : xmlFile {
    static $ns = "http://schemas.openxmlformats.org/package/2006/content-types"

    [bool] HasPngNode() {
        $pngnode = ( `
                $this.xdoc | `
                Select-Xml "/x:Types/x:Default[Extension='png']" -Namespace $this.nsm `
        )
        Return ($null -ne $pngnode )
    }

    [void] AddPngNode() {
        If ($this.HasPngNode()) {
            Return
        }
        
        $newNode = $this.xdoc.CreateElement("Default", [contentTypesXMLFile]::ns)
        $newNode.SetAttributeNode("Extension", "").Value = "png"
        $newNode.SetAttributeNode("ContentType", "").Value = "image/png"
        $this.xdoc.DocumentElement.AppendChild($newNode)
    }

    contentTypesXMLFile([string] $filePath) : base() {
        $this.setns([contentTypesXMLFile]::ns)
        # Content Types file has square brackets in the name
        # Literalpath _must_ be used.
        $this.xdoc = [xml] (Get-Content -LiteralPath $filePath -ErrorAction Stop)
        $this.ValidateSchema()
    }
}

Class relsXMLFile : xmlFile {
    static $ns = "http://schemas.openxmlformats.org/package/2006/relationships"

    [void] ValidateSchema() {
        $rootnode = $this.xdoc | Select-Xml "/x:Relationships" -Namespace $this.nsm
        If ($null -ne $rootnode) {
            $this.isvalid = $true
        }
        Else {
            $this.isvalid = $false
        }
    }

    [void] AddRel([string]$id, [string]$type, [string] $target) {
        $newNode = $this.xdoc.CreateElement("Relationship", $this.ns)
        $newNode.SetAttributeNode("Id", "").Value = $id
        $newNode.SetAttributeNode("Type", "").Value = $type
        $newNode.SetAttributeNode("Target", "").Value = $target
        
        $this.xdoc.DocumentElement.AppendChild($newNode)
    }

    [void] AddImageRel([string]$id, [string]$imagePath) {
        $this.AddRel( `
                $id, `
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", `
                $imagePath `
        )
    }

    [void] AddCustomUIRel([string] $id, [string]$customUIPath) {
        $this.AddRel( `
                $id, `
                "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility", `
                $customUIPath `
        )
    }

    relsXMLFile([string] $filePath) : base($filePath, [relsXMLFile]::ns ) {

    }

    relsXMLFile() : base() {
        $relsns = [relsXMLFile]::ns
        $this.setns($relsns)
        $this.xdoc = [xml]"<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n<Relationships xmlns='$relsns'></Relationships>"
        $this.isValid = $true
    }
}

Class customUIXMLFile : xmlFile {
    static $ns = "http://schemas.microsoft.com/office/2009/07/customui"

    [void] ValidateSchema() {
        $this.isValid = $false

        $ribbonElement = $this.xdoc | `
            Select-Xml  "/x:customUI/x:ribbon[@startFromScratch]" `
            -Namespace $this.nsm

        If ($null -eq $ribbonElement ) {
            Throw "Wrong schema found."
        }
    
        # Validate that there is only one ribbon element in the right place
        If ($ribbonElement.Count -ne 1) {
            Throw "There can be only one ribbon element."
        }
    
        # Validate that the ribbon element does not start from scratch
        If ($ribbonElement.Node.startFromScratch -ne "false") {
            Throw "Custom UIs are not allowed to have a from-scratch ribbon."
        }

        $this.isValid = $true
    }

    [string] InnerXml() {
        $this.CheckValidity()
        Return $this.xdoc.InnerXml
    }

    [object] GetTabs() {
        $this.CheckValidity()
        $tabs = $this.xdoc | Select-Xml `
            "//x:tabs/x:tab" `
            -Namespace $this.nsm
        If ($null -eq $tabs) {
            Return @()
        }
        Return $tabs
    }

    [object] GetButtons() {
        $this.CheckValidity()
        $buttons = $this.xdoc | Select-Xml `
            "/x:customUI/x:ribbon/x:tabs//x:button" `
            -Namespace $this.nsm
        If ($null -eq $buttons) {
            Return @()
        }
        Return $buttons
    }

    [bool] Merge([customUIXMLFile] $otherfile) {
        $somethingAdded = $false

        $otherfile.GetTabs() | ForEach-Object {
            If ($null -ne $_) {
                $addedForTab = $this.processTab($_.Node)
                $somethingAdded = $somethingAdded -or $addedForTab
            }
        }

        If ($somethingAdded) {
            $this.ValidateSchema()
        }

        Return $somethingAdded
    }

    [bool] hidden processElement($currentElement, $currentElementQuery, $parentElementQuery) {        
        $existingElement = $this.xdoc | `
            Select-Xml $currentElementQuery -Namespace $this.nsm
    
        If ($null -eq $existingElement) {
            $inode = $this.xdoc.ImportNode($currentElement, $true)
            $parentnode = ( `
                    $this.xdoc | `
                    Select-Xml $parentElementQuery -Namespace $this.nsm `
            ).Node
            $parentnode.AppendChild($inode)
    
            Return $true
        }
    
        Return $false
    }

    [bool] hidden processTab($currentTab) {
        $currentTabQuery = "//x:tabs/x:tab[@idMso='$($currentTab.idMso)']"
    
        $tabAdded = $this.processElement( `
                $currentTab, `
                $currentTabQuery, `
                "//x:ribbon/x:tabs" `
        )
        
        If (-not $tabAdded) {
            # If the tab is already there, iterate groups
            $groupsOrButtonsAdded = $false
            $groups = ( `
                    $currentTab | `
                    Select-Xml "//x:tab/x:group" -Namespace $this.nsm `
            )
            
            # Write-Verbose "Adding $($groups.Count) groups in tab $($currentTab.idMso)"
            $groups | ForEach-Object {
                $somethingAdded = $this.processTabGroup($_.Node, $currentTabQuery)
                $groupsOrButtonsAdded = $groupsOrButtonsAdded -or $somethingAdded
            }

            Return $groupsOrButtonsAdded
        }

        # Write-Verbose "Added tab $($currentTab.idMso)"
        Return $tabAdded
    }

    [bool] hidden processTabGroup($currentGroup, $currentTabQuery) {
        $currentGroupQuery = "$currentTabQuery/x:group[@id='$($currentGroup.id)']"
    
        $groupAdded = $this.processElement( `
                $currentGroup, `
                $currentGroupQuery, `
                $currentTabQuery `
        )
        
        If (-not $groupAdded) {
            # If group is already there, iterate buttons
            $buttonsAdded = $false
            $buttons = ( `
                    $currentGroup | `
                    Select-Xml "//x:group/x:button" -Namespace $this.nsm `
            )

            # Write-Verbose "Adding $($buttons.Count) buttons in group $($currentGroup.id)"
            $buttons | ForEach-Object {
                # Write-Verbose "Adding button $($_.Node.id)"
                $buttonAdded = $this.processButton($_.Node, $currentGroupQuery)
                $buttonsAdded = $buttonsAdded -or $buttonAdded
            }

            Return $buttonsAdded
        }

        # Write-Verbose "Added group $($currentGroup.id)"
        Return $groupAdded
    }

    [bool] hidden processButton($currentButton, $currentGroupQuery) {
        $currentButtonQuery = "$currentGroupQuery/x:button[@id='$($currentButton.id)']"
        
        $buttonAdded = $this.processElement( `
                $currentButton, `
                $currentButtonQuery, `
                $currentGroupQuery `
        )

        If (-not $buttonAdded) {
            # Current button should NOT exist in merged
            Throw "Duplicate button id: $($currentButton.id)"
        }

        # Write-Verbose "Added button $($currentButton.id)"
        Return $buttonAdded
    }

    customUIXMLFile([string] $filePath) : base($filePath, [customUIXMLFile]::ns ) {}

    customUIXMLFile() : base() {
        $this.setns([customUIXMLFile]::ns)
        $this.xdoc = [xml]'<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon startFromScratch="false"><tabs></tabs></ribbon></customUI>'
        $this.isValid = $true
    }
}

#endregion

#region MVPackage

Class MVPackage {
    [string] $Name
    [string] $Path

    [customUIXMLFile] hidden $cui
    [string] hidden $validationErr

    [string] CustomUIPath() {
        Return (Join-Path $this.Path "\CustomUI\customUI14.xml")
    }

    [string] CustomUIDir() {
        Return (Join-Path $this.Path "\CustomUI")
    }

    [bool] HasCustomUI() {
        Return (Test-Path $this.CustomUIPath())
    }

    [customUIXMLFile] CustomUI() {
        If (-not $this.HasCustomUI()) {
            Return $null
        }

        If ($null -eq $this.cui) {
            $this.cui = [customUIXMLFile]::New($this.CustomUIPath())
        }

        Return $this.cui
    }

    [void] Ensure() {
        [MVLibrary]::EnsureDirectory($this.Path)

        [MVLibrary]::EnsureDirectory((Join-Path $this.Path "Tests"))
    }

    [void] AddCustomUI() {
        If ($this.HasCustomUI()) {
            Throw "Package $($this.Name) already has custom UI."
        }

        [MVLibrary]::EnsureDirectory($this.CustomUIDir())
        
        $newcui = [customUIXMLFile]::New()
        $newcui.Save($this.CustomUIPath())

        $this.cui = $newcui
    }

    [bool] ValidateCustomUI() {
        $result = $true

        # Validate that all button image files are present 
        $currentcui = $this.CustomUI()
        $buttons = $currentcui.GetButtons()

        foreach ($item in $buttons) {
            $imagename = "$($item.Node.image)"
            If ("" -ne $imagename) {
                $imagepath = Join-Path $this.CustomUIDir() "$imagename.png"
                If (-not (Test-Path $imagepath)) {
                    $this.validationErr = "Image $imagepath not present in $($this.Name)."
                    $result = $false
                    break
                }
            }
        }

        Return $result
    }

    [string] ValidationError() {
        Return $this.validationErr
    }

    [void] Delete() {
        Remove-Item $this.Path -Force -Recurse -ErrorAction Stop
    }

    MVPackage($packageName, $packagePath) {
        $this.Name = $packageName
        $this.Path = $packagePath
    }
}

#endregion

#region MVBuilder

Class MVBuilder {
    [string] $Name
    [string] $Type
    [string] $OutPath

    # Will be overridden in application-specific builders
    [string] OfficeApplication() {
        Return "None"
    }

    # Will be overridden in application-specific builders
    [string] FileExtension([string] $buildType) {
        Return ".none"
    }

    [string] TargetPathNoExtension() {
        Return  Join-Path $this.OutPath $this.Name
    }

    [string] TargetPath() {
        $buildext = $this.FileExtension($this.Type)
        $targetPath = Join-Path $this.OutPath "$($this.Name).$buildext"
        Return $targetPath
    }

    [string] TargetUIDir() {
        Return (Join-Path $this.OutPath "$($this.Name).UI")
    }

    [void] Delete() {
        $buildfilepath = $this.TargetPath()
        $buildexpandeddir = "$buildfilepath.d"
        $builduidir = $this.TargetUIDir()
        Remove-Item $builduidir -Recurse -Force -ErrorAction Ignore
        Remove-Item $buildexpandeddir -Recurse -Force -ErrorAction Ignore
        Remove-Item $buildfilepath -Force -ErrorAction Ignore
    }

    [void] Build([MVPackage[]]$Packages) {
        If ($Packages.Count -eq 0) {
            Return
        }

        # Merge Office document
        $this.MergeOfficeDocument($Packages)

        # Merge Custom UI
        $cuiMerged = $this.MergeCustomUI($Packages)

        # Expand Office document, add Custom UI, and zip it
        # up again.
        If (-not(Test-Path $this.TargetPath() -PathType Leaf)) {
            Write-Host "Could not find built Office file. Skipping UI addition..."
            Return
        }

        If (-not $cuiMerged) {
            Write-Host "Skipping UI addition..."
            Return
        }

        $this.AddUIToDocument($this.TargetPath(), $this.TargetUIDir())
    }

    [void] MergeVBAModules([MVPackage[]]$Packages, $vbp) {
        $Packages | ForEach-Object {
            Write-Host "  Processing package $($_.Name)..."

            $packagepathspec = "$($_.Path)\*"
            $basfiles = (Get-ChildItem -Path $packagepathspec -Include *.bas, *.frm, *.cls)
            $basfiles | ForEach-Object {
                $newComponent = $vbp.VBComponents.Import($_.FullName)
                Write-Host "    Merged $($newComponent.Name)"     
            }
        }
    }

    # Will be overridden in application-specific builders
    [void] MergeOfficeDocument([MVPackage[]]$Packages) {
        Write-Host "Skipping merging Office $($this.Type)."
    }

    [bool] MergeCustomUI([MVPackage[]] $Packages) {
        Write-Host "Merging Custom UI..."        
        
        # Initialize Custom UI Directory and object
        $mergedCuiDir = $this.TargetUIDir()
        $mergedImageDir = (Join-Path $mergedCuiDir "images")
        Remove-Item $mergedCuiDir -Recurse -Force -ErrorAction Ignore
        $newcui = [customUIXMLFile]::New()
        $cuiMerged = $false
        
        foreach ($currentPackage in $Packages) {   
            # Merge Custom CUI 
            If ($currentPackage.HasCustomUI() ) {
                If (-not $currentPackage.ValidateCustomUI()) {
                    Throw "Package $($currentPackage.Name) could not be merged: $($currentPackage.validationErr)"
                }

                Write-Host "  Processing package $($currentPackage.Name)..."

                $packageCuiMerged = $newcui.Merge($currentPackage.CustomUI())
                $cuiMerged = $cuiMerged -or $packageCuiMerged
            
                If ($packageCuiMerged) {
                    [MVLibrary]::EnsureDirectory($mergedImageDir)
                    # Copy images
                    Copy-Item "$($currentPackage.Path)\CustomUI\*png" -Destination $mergedImageDir -Force
                }
            }
        }

        # If any Custom UI merging happened, save customUI
        # file and build rels file
        If ($cuiMerged) {
            [MVLibrary]::EnsureDirectory($mergedCuiDir)
            $newcuipath = Join-Path $this.TargetUIDir() "customUI14.xml"
            $newcui.Save($newcuipath)
            
            # Set up rels for copied images
            $imageCount = (Get-ChildItem $mergedImageDir -Filter "*.png" -ErrorAction Ignore).Count
            If ($imageCount -gt 0) {
                Write-Host "  Adding image rels file..."

                $relsdir = Join-Path $mergedCuiDir "_rels"
                $relsfilepath = Join-Path $relsdir "customUI14.xml.rels"
            
                [MVLibrary]::EnsureDirectory($relsdir)
                $relsfile = [relsXMLFile]::New()
            
                $newcui.GetButtons() | ForEach-Object {
                    $currentButton = $_.Node
                    $relsfile.AddImageRel($currentButton.image, "images/$($currentButton.image).png")
                }
            
                $relsfile.Save($relsfilepath)
            }

            Write-Host "Custom UI done."
        }
        Else {
            Write-Host "No custom UI present."
        }

        Return $cuiMerged
    }

    [void] AddUIToDocument([string]$officefilename, [string]$cuidir) {
        Write-Host "Adding UI to Office $($this.Type)..."
        $officezipfilename = "$officefilename.zip"
        $expanddir = "$officefilename.d"
    
        Remove-Item $officezipfilename -Force -ErrorAction Ignore
        Remove-Item $expanddir -Recurse -Force -ErrorAction Ignore
        
        # The PowerShell Expand-Archive cmdlet only works if the
        # file extension is .zip. So rename the office document,
        # expand to a directory, and delete the document
        Move-Item $officefilename $officezipfilename
        Expand-Archive $officezipfilename -DestinationPath $expanddir -ErrorAction Stop
        Remove-Item $officezipfilename -Force -ErrorAction Ignore
        
        # Move the previously consolidated custom UI directory
        # inside the expanded Office document directory
        Move-Item $cuiDir -Destination "$expanddir\customUI"
    
        # Edit the main rels file of the OpenOfficeXML document
        # to include the custom UI
        $relsFilePath = "$expanddir\_rels\.rels"
        $relsFile = [relsXMLFile]::New($relsFilePath)
        $newRelId = "R" + [string](Get-Random)
        $relsFile.AddCustomUIRel($newRelId, "/customUI/customUI14.xml")
        $relsFile.Save($relsFilePath)
    
        # Edit content-types file of the OpenOfficeXML document
        # to include png files if not already present.
        # Square brackets are tricky in PowerShell, so all operations
        # on this file will use the -LiteralPath parameter.
        $ctFilePath = "$expanddir\[Content_Types].xml"
        $ctFile = [contentTypesXMLFile]::New($ctFilePath)
        $ctFile.AddPngNode()
        $ctFile.Save($ctFilePath)
        
        # Re-compress the expanded Office document directory
        # with the custom UI files included, to a .zip file
        # because the Compress-Archive cmdlet will only zip to
        # that kind of file. Note that the _contents_ of the
        # directory are zipped. Office document files are zip
        # files with files and directories directly under the
        # root.    
        Compress-Archive "$expanddir\*" $officezipfilename -ErrorAction Stop
    
        # Rename it to the original name.
        Move-Item $officezipfilename $officefilename
    
        # Delete the expanded directory.
        Remove-Item $expanddir -Recurse -Force -ErrorAction Ignore
        
        Write-Host "Adding UI done."
    }


    MVBuilder([string] $buildName, [string] $buildType, [string] $buildParentDir) {
        $this.Name = $buildName
        $this.Type = $buildType
        $this.OutPath = $buildParentDir
    }
}

Class MVBuilderPPT : MVBuilder {
    static $BuildExtensions = @{
        "AddIn"    = "ppam";
        "Document" = "pptm"
    }

    # Override for PowerPoint
    [string] OfficeApplication() {
        Return "PowerPoint"
    }

    [string] FileExtension([string] $buildType) {
        Return [MVBuilderPPT]::BuildExtensions[$buildType]
    }

    [void] MergeOfficeDocument([MVPackage[]]$Packages) {
        Write-Host "Merging PowerPoint $($this.Type)..."

        $ppa = $null
        $newPpt = $null
        
        # Create PowerPoint object
        Try {
            $ppa = New-Object -ComObject PowerPoint.Application
        
            $newPpt = $ppa.Presentations.Add($false)
        }
        Catch {
            Throw "PowerPoint does not seem to be available."
        }
        
        Try {
            # Try to get the VBA project object.
            # If VBA Object model access is not trusted, $vbp
            # will contain $null
            $vbp = $newPpt.VBProject
            
            If ($null -eq $vbp) {
                Throw "Access to VBA Object Model not trusted. Please check the Trust Access to the VBA Object model checkbox in the PowerPoint Trust Centre."
            }
    
            $this.MergeVBAModules($Packages, $vbp)

            If ($this.Type -eq 'Document') {
                # 25 = ppSaveAsOpenXMLPresentationMacroEnabled
                $newPpt.SaveAs($this.TargetPathNoExtension(), 25, $false)
            }
            Else {
                # 30 = ppSaveAsOpenXMLAddin 
                $newPpt.SaveAs($this.TargetPathNoExtension(), 30, $false)
            }

            Write-Host "PowerPoint $($this.Type) done."
        }
        Finally {
            $newPpt.Close()
            $newPpt = $null
    
            $ppa.Quit()
            $ppa = $null
        }
    }

    MVBuilderPPT([string] $buildName, [string] $buildType, [string] $buildParentDir) :base($buildName, $buildType, $buildParentDir) {

    }
}

Class MVBuilderExcel : MVBuilder {
    static $BuildExtensions = @{
        "AddIn"    = "xlam";
        "Document" = "xlsm"
    }

    # Override for Excel
    [string] OfficeApplication() {
        Return "Excel"
    }

    [string] FileExtension([string] $buildType) {
        Return [MVBuilderExcel]::BuildExtensions[$buildType]
    }

    [void] MergeOfficeDocument([MVPackage[]]$Packages) {
        Write-Host "Merging Excel $($this.Type)..."

        $xla = $null
        $newWkbk = $null
        
        # Create Excel object
        Try {
            $xla = New-Object -ComObject Excel.Application
        
            $newWkbk = $xla.Workbooks.Add()
        }
        Catch {
            Throw "Excel does not seem to be available."
        }
        
        Try {
            # Try to get the VBA project object.
            # If VBA Object model is not trusted, $vbp
            # will contain $null
            $vbp = $newWkbk.VBProject
            
            If ($null -eq $vbp) {
                Throw "Access to VBA Object Model not trusted. Please check the Trust Access to the VBA Object model checkbox in the Macro Settings section of the Excel Trust Centre."
            }

            $this.MergeVBAModules($Packages, $vbp)
        
            If ($this.Type -eq 'Document') {
                # 52 = xlOpenXMLWorkbookMacroEnabled
                $newWkbk.SaveAs($this.TargetPathNoExtension(), 52)
            }
            Else {
                # 55 = xlOpenXMLAddIn 
                $newWkbk.SaveAs($this.TargetPathNoExtension(), 55)
            }

            Write-Host "Excel $($this.Type) done."
        }
        Finally {
            $newWkbk.Close()
            $newWkbk = $null
    
            $xla.Quit()
            $xla = $null
        }
    }

    MVBuilderExcel([string] $buildName, [string] $buildType, [string] $buildParentDir) :base($buildName, $buildType, $buildParentDir) {

    }
}

Class MVBuilderWord : MVBuilder {
    static $BuildExtensions = @{
        "AddIn"    = "";
        "Document" = "docm"
    }

    # Override for Excel
    [string] OfficeApplication() {
        Return "Word"
    }

    [string] FileExtension([string] $buildType) {
        Return [MVBuilderWord]::BuildExtensions[$buildType]
    }

    [void] MergeOfficeDocument([MVPackage[]]$Packages) {
        Write-Host "Merging Word $($this.Type)..."

        $wda = $null
        $newDoc = $null
        
        # Create Word object
        Try {
            $wda = New-Object -ComObject Word.Application
        
            $newDoc = $wda.Documents.Add()
        }
        Catch {
            Throw "Word does not seem to be available."
        }
        
        Try {
            # Try to get the VBA project object.
            # If VBA Object model is not trusted, $vbp
            # will contain $null
            $vbp = $newDoc.VBProject
            
            If ($null -eq $vbp) {
                Throw "Access to VBA Object Model not trusted. Please check the Trust Access to the VBA Object model checkbox in the Word Trust Centre."
            }

            $this.MergeVBAModules($Packages, $vbp)
        
            If ($this.Type -eq 'Document') {
                # 13 = wdFormatXMLDocumentMacroEnabled
                $newDoc.SaveAs2($this.TargetPathNoExtension(), 13)
            }
            Else {
                # 15 = wdFormatXMLTemplateMacroEnabled 
                $newDoc.SaveAs2($this.TargetPathNoExtension(), 15)
            }

            Write-Host "Word $($this.Type) done."
        }
        Finally {
            $newDoc.Close()
            $newDoc = $null
    
            $wda.Quit()
            $wda = $null
        }
    }

    MVBuilderWord([string] $buildName, [string] $buildType, [string] $buildParentDir) :base($buildName, $buildType, $buildParentDir) {

    }
}
#endregion

#region MVLibrary

Class MVLibrary {
    [string[]] static $SupportedApplications = "none", "powerpoint", "excel", "word"
    [string] $Name
    [string] $Application

    [void] static EnsureDirectory([string]$OutPath) {
        If ("" -eq $OutPath) {
            Throw "Empty Path."
        }
    
        $outdirexists = (Test-Path -PathType Container $OutPath)
        If ($outdirexists -eq $false) {
            New-Item -ItemType Directory -Force -Path $OutPath -ErrorAction Stop
        }
    }

    [System.Management.Automation.ErrorRecord] static CmdLetError($err) {
        If ($_ -is [System.Management.Automation.ErrorRecord]) {
            return $err
        }

        return $PSCmdlet.WriteError( `
                [System.Management.Automation.ErrorRecord]::New( `
                    $err, `
                    "MOVBA.Error", `
                    [System.Management.Automation.ErrorCategory]::NotSpecified, `
                    $null `
            ) `
        )
    }
    
    [void] static ThrowDirectoryException() {
        Throw "This CmdLet can only be used in a project root directory."
    }

    [MVLibrary] static GetProjectConfig() {       
        # Check if  the packages subdirectory is present
        If (-not ( `
                (Test-Path "$pwd\packages" -PathType Container ) `
                    -and 
                (Test-Path "$pwd\library.json" -PathType Leaf)
            )) {
            [MVLibrary]::ThrowDirectoryException()
        }

        $appconfig = ConvertFrom-Json (Get-Content -Raw "$pwd\library.json" -ErrorAction Ignore)
        If ($null -eq $appconfig) {
            [MVLibrary]::ThrowDirectoryException()
        }

        $lib = [MVLibrary]::New($appconfig.Name, $appconfig.Application)

        Return $lib
    }

    [void] static TestCurrentDirectory() {
        If ($null -eq [MVLibrary]::GetProjectConfig()) {
            [MVLibrary]::ThrowDirectoryException()
        }
    }

    [bool] static TestOfficeApplication([string] $Application) {
        If (-not ( `
                    [MVLibrary]::SupportedApplications.Contains( `
                        $Application.ToLowerInvariant() `
                ) `
            ) `
        ) {
            Throw "Application $Application not currently supported."
        }

        Return $true
    }

    [MVBuilder] static GetBuilder([string] $Name, [string] $Type, [string] $Application) {
        $result = $null
        Switch ($Application) {
            "none" { 
                [MVLibrary]::EnsureDirectory("$pwd\out")
                $result = [MVBuilder]::New($Name, $Type, "$pwd\out")
                Break
            }
            "powerpoint" { 
                [MVLibrary]::EnsureDirectory("$pwd\out")
                $result = [MVBuilderPPT]::New($Name, $Type, "$pwd\out")
                Break
            }
            "excel" {
                [MVLibrary]::EnsureDirectory("$pwd\out")
                $result = [MVBuilderExcel]::New($Name, $Type, "$pwd\out")
                Break
            }
            "word" {
                [MVLibrary]::EnsureDirectory("$pwd\out")
                $result = [MVBuilderWord]::New($Name, $Type, "$pwd\out")
                Break
            }
            Default { 
                Throw "Application $Application not currently supported."
            }
        }
        Return $result
    }

    [MVLibrary] static Init([string] $Name, [string] $Application) {
        [MVLibrary]::TestOfficeApplication($Application)

        $projectDir = (Join-Path $pwd $Name) 
        If (Test-Path $projectDir) {
            Throw "Directory '$Name' already exists."
        }

        New-Item $projectDir -ItemType Directory

        Push-Location $projectDir

        $appconfig = [PSCustomObject] @{"Name" = $Name; "Application" = $Application }
        $appconfigJSON = ConvertTo-Json $appconfig
        Set-Content -Value $appconfigJSON -Path "$pwd\library.json" -ErrorAction Stop

        # Create packages subdirectory
        If (-not (Test-Path "$pwd\packages" -PathType Container)) {
            New-Item "$pwd\packages" -ItemType Directory
        }

        # Create out subdirectory
        If (-not (Test-Path "$pwd\out" -PathType Container)) {
            New-Item "$pwd\out" -ItemType Directory
        }

        If (-not (Test-Path "$pwd\.gitignore" -PathType Leaf)) {
            Set-Content "$pwd\.gitignore" (@(
                    "# Office temporary files"
                    "*.tmp"
                    "~$*.ppt*"
                    "~$*.doc*"
                    "~$*.xls*"
                    ""
                    "# output directories"
                    "out/"
                ) -join "`n") 
        }

        $result = [MVLibrary]::New($appconfig.Name, $appconfig.Application)

        Write-Host "$($appconfig.Application) project initialised."
        Pop-Location

        return $result
    }

    [void] static Build([string] $Name, [string] $Type, [MVPackage[]]$Packages) {
        $appconfig = [MVLibrary]::GetProjectConfig()
        If ($null -eq $appconfig) {
            [MVLibrary]::ThrowDirectoryException()
        }

        [MVLibrary]::TestOfficeApplication($appconfig.Application)

        $builder = [MVLibrary]::GetBuilder($Name, $Type, $appconfig.Application)

        $builder.Build($Packages)
    }

    [MVPackage[]] static GetPackages([string[]] $PackageNames) {
        [MVLibrary]::TestCurrentDirectory()

        $packageDirQuery = Get-ChildItem -Path "$pwd\packages" -Directory
        If ($PackageNames.Count -gt 0) {
            $pnfc = $PackageNames | ForEach-Object { $_.ToLowerInvariant() }
            $packageDirQuery = $packageDirQuery | Where-Object { $pnfc.Contains($_.Name.ToLowerInvariant()) }
        }

        $packageDirQuery = $packageDirQuery | ForEach-Object { [MVPackage]::New($_.Name, $_.FullName) }

        Return $packageDirQuery
    }

    [MVPackage[]] static GetAllPackages() {
        Return [MVLibrary]::GetPackages(@())
    }

    [MVPackage] static GetPackage([string] $PackageName) {
        [MVLibrary]::TestCurrentDirectory()

        $packageDirQuery = Get-ChildItem -Path "$pwd\packages" -Directory | `
            Where-Object { $_.Name -eq $PackageName }

        $result = $packageDirQuery | ForEach-Object { [MVPackage]::New($_.Name, $_.FullName) }

        Return $result
    }

    [MVPackage] static NewPackage([string]$PackageName, [bool]$HasCustomUI) {
        [MVLibrary]::TestCurrentDirectory()

        $oldpackage = [MVLibrary]::GetPackage($PackageName)
        If ($null -ne $oldpackage) {
            Throw "Package $PackageName already exists."
        }

        # TODO: Consider switching to PS 6 minimum
        $newPackagePath = Join-Path (Join-Path "$pwd" "packages") "$PackageName"
        $newPackage = [MVPackage]::New($PackageName, $newPackagePath)
        $newPackage.Ensure()
        If ($HasCustomUI) {
            $newPackage.AddCustomUI()
        }

        Return $newPackage
    }

    [void] static RemovePackage([string]$PackageName) {
        [MVLibrary]::TestCurrentDirectory()

        $oldpackage = [MVLibrary]::GetPackage($PackageName)
        If ($null -eq $oldPackage) {
            Throw "Package $PackageName does not exist."
        }
    
        $oldPackage.Delete()    
    }

    MVLibrary([string] $libName, [string] $libApplication) {
        $this.Name = $libName
        $this.Application = $libApplication
    }
}

#endregion