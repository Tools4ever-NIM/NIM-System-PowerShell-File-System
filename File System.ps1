#
# File System.ps1 - IDM System PowerShell Script for File System integration.
#
# Any IDM System PowerShell Script is dot-sourced in a separate PowerShell context, after
# dot-sourcing the IDM Generic PowerShell Script '../Generic.ps1'.
#


$NrOfAccessProfiles = 3


#
# System functions
#

function Idm-SystemInfo {
    param (
        # Operations
        [switch] $Connection,
        [switch] $TestConnection,
        [switch] $Configuration,
        # Parameters
        [string] $ConnectionParams
    )

    Log info "-Connection=$Connection -TestConnection=$TestConnection -Configuration=$Configuration -ConnectionParams='$ConnectionParams'"
    
    if ($Connection) {
        @(
            @{
                name = 'paths_spec'
                type = 'textbox'
                label = 'Paths'
                tooltip = "Paths to collect. Separate multiple paths by '|'. Optionally suffix path with ':<n>' to collect 'n' levels deep."
                value = ''
            }
            @{
                name = 'excludes'
                type = 'textbox'
                label = 'Excludes'
                tooltip = "File name patterns to exclude. Separate multiple patterns by '|'. E.g. *\example excludes all folders with the name 'example' and their contents; *\example\* excludes the contents of all folders with the name 'example', not the folder itself."
                value = ''
            }
             @{
                name = 'recursive'
                type = 'checkbox'
                label = 'Recursive'
                value = $true
            }
			@{
                name = 'recursion_depth'
                type = 'textbox'
                label = 'Recursion Depth'
                tooltip = 'Max. depth of recursion'
                value = 1
                hidden = '!recursive'
            }
			
            @{
                name = 'ignoreACEPermissionErrors'
                type = 'checkbox'
                label = 'Ignore ACE Permission Errors'
                value = $false
            }
            @{
                name = 'skipFolderACL'
                type = 'checkbox'
                label = 'Skip Folder ACL''s'
                value = $false
            }
        )
    }

    if ($TestConnection) {
        $connection_params = ConvertSystemParams $ConnectionParams

        foreach ($path_spec in $connection_params.paths_spec) {
            Get-ChildItem -Force -LiteralPath $path_spec.path >$null
        }
    }

    if ($Configuration) {
        @(
            @{
                name = "principal_type"
                type = 'radio'
                label = 'Principal type'
                table = @{
                    rows = @(
                        @{ id = 'SecurityIdentifier'; text = 'SID (recommended)' }
                        @{ id = 'NTAccount';          text = 'Account name' }
                    )
                    settings_radio = @{
                        value_column = 'id'
                        display_column = 'text'
                    }
                }
                value = 'SecurityIdentifier'
            }
            @{
                name = 'separator'
                type = 'text'
                text = ''
            }

            foreach ($nr in 1..$Global:NrOfAccessProfiles) {
                $prefix = "access_profile_$($nr)_"

                @{
                    name = "$($prefix)enable"
                    type = 'checkbox'
                    label = "Access Profile $nr"
                    value = $nr -eq 1
                }
                @{
                    name = "$($prefix)property_name"
                    type = 'textbox'
                    label = 'Property name'
                    label_indent = $true
                    value = "AccessProfile$nr"
                    hidden = "!!"   # Permanently hidden
                }
                @{
                    name = "$($prefix)type"
                    type = 'combo'
                    label = 'Type'
                    label_indent = $true
                    table = @{
                        rows = @(
                            @{ id = [System.Security.AccessControl.AccessControlType]::Allow; text = 'Allow' }
                            @{ id = [System.Security.AccessControl.AccessControlType]::Deny;  text = 'Deny'  }
                        )
                        settings_combo = @{
                            value_column = 'id'
                            display_column = 'text'
                        }
                    }
                    value = [System.Security.AccessControl.AccessControlType]::Allow
                    hidden = "!$($prefix)enable"
                }
                @{
                    name = "$($prefix)permissions"
                    type = 'checkgroup'
                    label = 'Permissions'
                    label_indent = $true
                    table = @{
                        rows = @(
                            @{ id = [String][Uint32][System.Security.AccessControl.FileSystemRights]::FullControl;    text = 'Full control'   }
                            @{ id = [String][Uint32][System.Security.AccessControl.FileSystemRights]::Modify;         text = 'Modify'         }
                            @{ id = [String][Uint32][System.Security.AccessControl.FileSystemRights]::ReadAndExecute; text = 'Read & execute' }
                            @{ id = [String][Uint32][System.Security.AccessControl.FileSystemRights]::Read;           text = 'Read'           }
                            @{ id = [String][Uint32][System.Security.AccessControl.FileSystemRights]::Write;          text = 'Write'          }
                        )
                        settings_checkgroup = @{
                            value_column = 'id'
                            display_column = 'text'
                        }
                    }
                    value = @([String][Uint32][System.Security.AccessControl.FileSystemRights]::Modify)
                    hidden = "!$($prefix)enable"
                }
                @{
                    name = "$($prefix)permissions_matching"
                    type = 'radio'
                    label = 'Permissions matching'
                    label_indent = $true
                    table = @{
                        rows = @(
                            @{ id = 'effective'; text = 'Effective (recommended) - subsequent explicit ACEs accumulate to at least these permissions (not considering group memberships)' }
                            @{ id = 'exact';     text = 'Exact - at least one explicit ACE exactly has these permissions' }
                        )
                        settings_radio = @{
                            value_column = 'id'
                            display_column = 'text'
                        }
                    }
                    value = 'effective'
                    hidden = "!$($prefix)enable"
                }
                @{
                    name = "$($prefix)applies_to"
                    type = 'combo'
                    label = 'Applies to'
                    label_indent = $true
                    table = @{
                        rows = @(
                            # Mapping, see: https://docs.microsoft.com/en-us/previous-versions/dotnet/netframework-4.0/ms229747(v=vs.100)
                            @{ id = '';       text = 'This folder only'                  }
                            @{ id = 'OICI';   text = 'This folder, subfolders and files' }
                            @{ id = 'CI';     text = 'This folder and subfolders'        }
                            @{ id = 'OI';     text = 'This folder and files'             }
                            @{ id = 'OICIIO'; text = 'Subfolders and files only'         }
                            @{ id = 'CIIO';   text = 'Subfolders only'                   }
                            @{ id = 'OIIO';   text = 'Files only'                        }
                        )
                        settings_combo = @{
                            value_column = 'id'
                            display_column = 'text'
                        }
                    }
                    value = 'OICI'
                    hidden = "!$($prefix)enable"
                }
                @{
                    name = 'separator'
                    type = 'text'
                    text = ''
                }
            }

            @{
                name = 'separator'
                type = 'text'
                text = '*** ExplicitACEs TABLE BELOW IS FOR INFORMATIONAL PURPOSES ONLY ***'
            }
        )
    }

    Log info "Done"
}


#
# CRUD functions
#

$Properties = @{
    Folder = @(
        @{ name = 'FullName';          default = $true; key = $true }
        @{ name = 'Attributes';                                     }
        @{ name = 'CreationTime';                                   }
        @{ name = 'Depth';                                          }
        @{ name = 'InheritanceEnable'; default = $true;             }
        @{ name = 'LastAccessTime';                                 }
        @{ name = 'LastWriteTime';                                  }
        @{ name = 'Name';              default = $true;             }
        @{ name = 'Owner';             default = $true;             }
        @{ name = 'Path';                                           }
    )
}


# Default properties and IDM properties are the same
foreach ($key in $Properties.Keys) {
    for ($i = 0; $i -lt $Properties.$key.Count; $i++) {
        if ($Properties.$key[$i].default) {
            $Properties.$key[$i].idm = $true
        }
    }
}


function Idm-FolderCreate {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        $system_params = ConvertSystemParams $SystemParams

        @{
            semantics = 'create'
            parameters = @(
                @{ name = 'FullName';          allowance = 'mandatory' }
                @{ name = 'InheritanceEnable'; allowance = 'optional'  }
               #@{ name = 'Owner';             allowance = 'optional'  }

                foreach ($nr in 1..$Global:NrOfAccessProfiles) {
                    $prefix = "access_profile_$($nr)_"

                    if ($system_params["$($prefix)enable"]) {
                        @{ name = $system_params["$($prefix)property_name"]; allowance = 'optional' }
                    }
                }

                @{ name = '*'; allowance = 'prohibited' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertSystemParams $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        LogIO info "New-Item" -In -ItemType 'Directory' -Path $function_params.FullName
            $rv = New-Item -ItemType Directory -Path $function_params.FullName | Select-Object -Property 'FullName'
        LogIO info "New-Item" -Out $rv

        ModifyFileSecurityDescriptor $system_params $function_params $function_params.FullName | Out-Null

        $rv
    }

    Log info "Done"
}


function Idm-ExplicitACEsRead {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        # Purposely left empty: nothing to configure
    }
    else {
        #
        # Execute function
        #

        $system_params = ConvertSystemParams $SystemParams

        $out = New-Object System.Collections.ArrayList

        foreach ($path_spec in $system_params.paths_spec) {
            Log debug "Path: $($path_spec.path)"
            $gci_args = @{
                Directory   = $true
                Force       = $true
                LiteralPath = $path_spec.path
                Recurse     = $system_params.recursive
                Depth 	    = 0
                ErrorAction = 'SilentlyContinue'
            }
			
            if($system_params.recursive) {
                $gci_args.Depth = $system_params.recursion_depth
            }

            if ($path_spec.depth -ge 0) {
                $gci_args.Depth = $path_spec.depth
            }

            try {
                # This is to correct error messages, e.g.:
                #   "Cannot find drive. A drive with the name 'x' does not exist" instead of
                #   "A parameter cannot be found that matches parameter name 'Directory'".
                Get-ChildItem -Force -LiteralPath $path_spec.path >$null

                # For directories, Get-ChildItem returns [System.IO.DirectoryInfo]
                Get-ChildItem @gci_args | ForEach-Object {
                    foreach ($exclude in $system_params.excludes) {
                        if ($_.FullName -ilike $exclude) { return }
                    }

                    $_
                } | ForEach-Object {
                    $full_name = $_.FullName

                    # For directories, GetAccessControl() returns [System.Security.AccessControl.DirectorySecurity],
                    # which is the same as Get-Acl returns.
                    $sd = $_.GetAccessControl()
                    # GetAccessRules() returns [System.Security.AccessControl.FileSystemAccessRule]
                    $acl = $sd.GetAccessRules($true, $false, $system_params.principal_type)    # includeExplicit, !includeInherited
                    $ix = 0

                    $acl | ForEach-Object {
                        $ht = [ordered]@{
                            FullName          = $full_name
                            Ix                = $ix
                            AccessControlType = $_.AccessControlType
                            IdentityReference = $_.IdentityReference
                            FileSystemRights  = $_.FileSystemRights
                            InheritanceFlags  = $_.InheritanceFlags
                            PropagationFlags  = $_.PropagationFlags
                        }

                        $out.Add((New-Object -TypeName PSObject -Property $ht)) | Out-Null
                        $ix++
                    }
                }
            }
            catch {
                if($system_params.ignoreACEPermissionErrors)
                {
                    Log warning "Failed: $_"
                    Write-Warning $_
                } else {
                    Log error "Failed: $_"
                    Write-Error $_
                }
            }
        }

        $out | Sort-Object { $_.FullName; $_.Ix }
    }

    Log info "Done"
}


function Idm-FoldersRead {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        $system_params = ConvertSystemParams $SystemParams

        $all_properties = $Global:Properties.Folder + @(
            foreach ($nr in 1..$Global:NrOfAccessProfiles) {
                $prefix = "access_profile_$($nr)_"

                if ($system_params["$($prefix)enable"]) {
                    @{ name = $system_params["$($prefix)property_name"]; default = $nr -eq 1; idm = $nr -eq 1 }
                }
            }
        )

        @(
            @{
                name = 'properties'
                type = 'grid'
                label = 'Properties'
                table = @{
                    rows = @( $all_properties | ForEach-Object {
                        @{
                            name = $_.name
                            usage_hint = @( @(
                                foreach ($key in $_.Keys) {
                                    if ($key -eq 'idm') {
                                        $key.Toupper()
                                    }
                                    elseif ($key -ne 'name') {
                                        $key.Substring(0,1).Toupper() + $key.Substring(1)
                                    }
                                }
                            ) | Sort-Object) -join ' | '
                        }
                    })
                    settings_grid = @{
                        selection = 'multiple'
                        key_column = 'name'
                        checkbox = $true
                        filter = $true
                        columns = @(
                            @{
                                name = 'name'
                                display_name = 'Name'
                            }
                            @{
                                name = 'usage_hint'
                                display_name = 'Usage hint'
                            }
                        )
                    }
                }
                value = ($all_properties | Where-Object { $_.default }).name
            }
        )

    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertSystemParams $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        if ($function_params.properties.length -eq 0) {
            # No properties selected: select defaults

            $all_properties = $Global:Properties.Folder + @(
                foreach ($nr in 1..$Global:NrOfAccessProfiles) {
                    $prefix = "access_profile_$($nr)_"

                    if ($system_params["$($prefix)enable"]) {
                        @{ name = $system_params["$($prefix)property_name"]; default = $nr -eq 1; idm = $nr -eq 1 }
                    }
                }
            )

            $function_params.properties = ($all_properties | Where-Object { $_.default }).name
        }

        # Assure key is the first column
        $key = ($Global:Properties.Folder | Where-Object { $_.key }).name
        $function_params.properties = @($key) + @($function_params.properties | Where-Object { $_ -ne $key })

        $access_profiles = GetAccessProfiles $system_params $function_params

        foreach ($path_spec in $system_params.paths_spec) {
			Log debug "Path: $($path_spec.path)"
            $path_with_backslash = AppendBackslashToPath $path_spec.path

            $gci_args = @{
                Path 		= $path_spec.path
                Depth       = 0
				Exclude 	= $system_params.excludes
            }

            if($system_params.recursive) {
				Log debug "Setting depth [$($system_params.recursion_depth)]"
                $gci_args.Depth = $system_params.recursion_depth
            }

            if ($path_spec.depth -ge 0) {
                $gci_args.Depth = $path_spec.depth
            }

            LogIO info 'GetItemsWithDepth' -In @gci_args -Properties $function_params.properties

            try {
				# For directories, Get-ChildItem returns [System.IO.DirectoryInfo]
				GetItemsWithDepth @gci_args | ForEach-Object {
                    foreach ($exclude in $system_params.excludes) {
                        if ($_.FullName -ilike $exclude) { return }
                    }
                    $_
                } | ForEach-Object {
					# For directories, GetAccessControl() returns [System.Security.AccessControl.DirectorySecurity],
                    # which is the same as Get-Acl returns.
                    if($system_params.skipFolderACL) {
                        $ht = @{
                            Attributes        = ($_.Attributes -split ',' | ForEach-Object { $h = $_; if ($h.Length -gt 0) { $h.Substring(0,1).Toupper() } }) -join ''
                            Depth             = $_.FullName.Substring($path_with_backslash.length).Split('\').Count - 1
                            InheritanceEnable = ''
                            Owner             = ''
                            Path              = $_.FullName.Substring(0, $_.FullName.length - $_.Name.Length)
                        }
                    } else {
                        $sd = $_.GetAccessControl()
                        
                        $ht = @{
                            Attributes        = ($_.Attributes -split ',' | ForEach-Object { $h = $_; if ($h.Length -gt 0) { $h.Substring(0,1).Toupper() } }) -join ''
                            Depth             = $_.FullName.Substring($path_with_backslash.length).Split('\').Count - 1
                            InheritanceEnable = $sd.AreAccessRulesProtected -eq $false
                            Owner             = $sd.GetOwner($system_params.principal_type).Value
                            Path              = $_.FullName.Substring(0, $_.FullName.length - $_.Name.Length)
                        }

                        $ht += GetIdentityReferencesMatchingAccessProfiles $sd $access_profiles $system_params.principal_type
                    }
                    
                    $_ | Add-Member -PassThru -Force -NotePropertyMembers $ht
                } | Select-Object $function_params.properties | Sort-Object { $_.FullName }
            }
            catch {
                Log error "Failed: $_"
                Write-Error $_
            }
        }

    }

    Log info "Done"
}


function Idm-FolderUpdate {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        $system_params = ConvertSystemParams $SystemParams

        @{
            semantics = 'update'
            parameters = @(
                @{ name = 'FullName';          allowance = 'mandatory' }
                @{ name = 'Name';              allowance = 'mandatory'  }

                foreach ($nr in 1..$Global:NrOfAccessProfiles) {
                    $prefix = "access_profile_$($nr)_"

                    if ($system_params["$($prefix)enable"]) {
                        @{ name = $system_params["$($prefix)property_name"]; allowance = 'optional' }
                    }
                }

                @{ name = '*'; allowance = 'prohibited' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params   = ConvertSystemParams $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        $full_name = $function_params.FullName

		LogIO info "Rename-Item" -In -LiteralPath $function_params.FullName -Destination $function_params.Name
			$rv = Rename-Item -PassThru -LiteralPath $function_params.FullName -NewName $function_params.Name | Select-Object -Property @{ Name = 'FullName'; Expression = {$_.FullName.TrimEnd('\')} }
		LogIO info "Rename-Item" -Out $rv

	    $function_params
    }

    Log info "Done"
}


function Idm-FolderDelete {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    if ($GetMeta) {
        #
        # Get meta data
        #

        @{
            semantics = 'delete'
            parameters = @(
                @{ name = 'FullName'; allowance = 'mandatory'  }
                @{ name = '*';        allowance = 'prohibited' }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $function_params = ConvertFrom-Json2 $FunctionParams

        LogIO info "Remove-Item" -In -Recurse $true -Force $true -LiteralPath $function_params.FullName -Confirm:$false
            $rv = Remove-Item -Recurse -Force -LiteralPath $function_params.FullName -Confirm:$false
        LogIO info "Remove-Item" -Out $rv

        $rv
    }

    Log info "Done"
}


#
# Helper functions
#

function ConvertSystemParams {
    param (
        [string] $InputParams
    )

    $params = ConvertFrom-Json2 $InputParams

    $params.paths_spec = @(
        $params.paths_spec.Split('|') | ForEach-Object {
            $value = $_
            if ($value.length -eq 0) { return }

            $p = $value.LastIndexOf(':')

            if ($p -le 1 -or ($p -eq 5 -and $value.IndexOf('\\?\') -eq 0)) {
                # No depth specified or part of drive letter
                @{
                    path  = $value
                    depth = -1
                }
            }
            else {
                @{
                    path  = $value.Substring(0, $p)
                    depth = $value.Substring($p + 1)
                }
            }
        }
    )

    $params.excludes = @(
        $params.excludes.Split('|') | ForEach-Object {
            $value = $_
            if ($value.length -eq 0) { return }

            $value
            $value + '\*'    # Probably always wanted
        }
    )

    $params.principal_type = if ($params.principal_type -eq 'NTAccount') { [System.Security.Principal.NTAccount] } else { [System.Security.Principal.SecurityIdentifier] }

    return $params
}


function AppendBackslashToPath {
    param (
        [string] $Path
    )

    if ($Path.length -eq 0 -or $Path.Substring($Path.length - 1) -eq ':') {
        # Do not append backslash, as it would result in an absolute path
        $Path
    }
    elseif ($Path.Substring($Path.length - 1) -eq '\') {
        # Already ends with a backslash
        $Path
    }
    else {
        $Path + '\'
    }
}


function GetAccessProfiles {
    param (
        [hashtable] $SystemParams,
        [hashtable] $FunctionParams
    )

    $is_read_call = $FunctionParams.ContainsKey('properties')

    @(
        foreach ($nr in 1..$Global:NrOfAccessProfiles) {
            $prefix = "access_profile_$($nr)_"

            if ($SystemParams["$($prefix)enable"]) {
                $property_name = $SystemParams["$($prefix)property_name"]

                if ( ( $is_read_call -and $FunctionParams.properties.Contains($property_name)) -or
                     (!$is_read_call -and $FunctionParams.ContainsKey($property_name)) ) {
                    @{
                        property_name     = $property_name
                        accessControlType = [System.Security.AccessControl.AccessControlType]$SystemParams["$($prefix)type"]
                        fileSystemRights  = [System.Security.AccessControl.FileSystemRights]$($fsr = [Uint32][System.Security.AccessControl.FileSystemRights]::Synchronize; $SystemParams["$($prefix)permissions"] | ForEach-Object { $fsr = $fsr -bor [Uint32]$_ }; $fsr)
                        exact_matching    = $SystemParams["$($prefix)permissions_matching"] -eq 'exact'
                        inheritanceFlags  = $(if ($SystemParams["$($prefix)applies_to"].IndexOf('OI') -ge 0) { [System.Security.AccessControl.InheritanceFlags]::ObjectInherit    } else { 0 }) -bor 
                                            $(if ($SystemParams["$($prefix)applies_to"].IndexOf('CI') -ge 0) { [System.Security.AccessControl.InheritanceFlags]::ContainerInherit } else { 0 })
                        propagationFlags  =   if ($SystemParams["$($prefix)applies_to"].IndexOf('IO') -ge 0) { [System.Security.AccessControl.PropagationFlags]::InheritOnly      } else { 0 }
                    }
                }
            }
        }
    )
}


function Condense-ACL {
    param (
        [System.Security.AccessControl.FileSystemAccessRule[]] $Acl
    )

    $condensed_acl = New-Object System.Collections.ArrayList

    $Acl | ForEach-Object {
        $fsr = $_.FileSystemRights

        # Convert Generic Rights to File System Rights
        $fsr = switch ($_.FileSystemRights -band -bnot [System.Security.AccessControl.FileSystemRights]::Delete) {
            0x10000000 { [System.Security.AccessControl.FileSystemRights]::FullControl; break }
            0x40000000 { [System.Security.AccessControl.FileSystemRights]::Write -bor [System.Security.AccessControl.FileSystemRights]::ReadPermissions -bor [System.Security.AccessControl.FileSystemRights]::Synchronize; break }
            0x80000000 { [System.Security.AccessControl.FileSystemRights]::Read -bor [System.Security.AccessControl.FileSystemRights]::Synchronize; break }
            0xA0000000 { [System.Security.AccessControl.FileSystemRights]::ReadAndExecute -bor [System.Security.AccessControl.FileSystemRights]::Synchronize; break }
            0xE0000000 { [System.Security.AccessControl.FileSystemRights]::Modify -bor [System.Security.AccessControl.FileSystemRights]::Synchronize; break }
            default    { $fsr }
        }

        for ($ix = $condensed_acl.Count-1; $ix -ge 0; $ix--) {
            if ($condensed_acl[$ix].IdentityReference -ne $_.IdentityReference) {
                # Different IdentityReference: irrelevant ACE
                continue
            }

            if ($condensed_acl[$ix].AccessControlType -ne $_.AccessControlType) {
                # Different AccessControlType: append this ACE
                $ix = -1
                break
            }

            if ($condensed_acl[$ix].FileSystemRights -ne $fsr) {
                # Different FileSystemRights: cannot condense
                continue
            }

            # Condense InheritanceFlags and PropagationFlags
            $condensed_acl[$ix].InheritanceFlags = [System.Security.AccessControl.InheritanceFlags]($condensed_acl[$ix].InheritanceFlags -bor  $_.InheritanceFlags)
            $condensed_acl[$ix].PropagationFlags = [System.Security.AccessControl.PropagationFlags]($condensed_acl[$ix].PropagationFlags -band $_.PropagationFlags)
            break
        }

        if ($ix -eq -1) {
            # Append this ACE
            $ace = @{
                FileSystemRights  = $fsr
                AccessControlType = $_.AccessControlType
                IdentityReference = $_.IdentityReference
                InheritanceFlags  = $_.InheritanceFlags
                PropagationFlags  = $_.PropagationFlags
            }

            $condensed_acl.Add((New-Object -TypeName PSObject -Property $ace)) | Out-Null
        }
    }

    $condensed_acl
}


function GetIdentityReferencesMatchingAccessProfiles {
    param (
        $SecDesc,
        $AccessProfiles,
        $PrincipalType
    )

    $ht = @{}

    # GetAccessRules() returns [System.Security.AccessControl.FileSystemAccessRule]
    $acl = Condense-ACL ($SecDesc.GetAccessRules($true, $false, $PrincipalType))    # includeExplicit, !includeInherited

    foreach ($ap in $AccessProfiles) {
        $ht[$ap.property_name] = @(
            if ($ap.exact_matching) {
                $acl | Where-Object {
                    $_.AccessControlType -eq $ap.accessControlType -and
                    ($_.FileSystemRights -bor [System.Security.AccessControl.FileSystemRights]::Synchronize) -eq $ap.fileSystemRights -and
                    $_.InheritanceFlags  -eq $ap.inheritanceFlags  -and
                    ($_.PropagationFlags -band -bnot [System.Security.AccessControl.PropagationFlags]::NoPropagateInherit) -eq $ap.propagationFlags
                } | ForEach-Object {
                    $_.IdentityReference.Value
                }
            }
            else {
                # Effective matching (not considering group memberships)
                @(
                    $acl | Where-Object {
                        $_.InheritanceFlags  -eq $ap.inheritanceFlags -and
                        ($_.PropagationFlags -band -bnot [System.Security.AccessControl.PropagationFlags]::NoPropagateInherit) -eq $ap.propagationFlags
                    } | ForEach-Object {
                        $_.IdentityReference
                    }
                ) | Select-Object -Unique | ForEach-Object {
                    $ir = $_
                    $mask = $ap.fileSystemRights

                    $acl | Where-Object {
                        $_.InheritanceFlags  -eq $ap.inheritanceFlags -and
                        ($_.PropagationFlags -band -bnot [System.Security.AccessControl.PropagationFlags]::NoPropagateInherit) -eq $ap.propagationFlags -and
                        $_.IdentityReference -eq $ir
                    } | ForEach-Object {
                        if ($_.AccessControlType -eq $ap.accessControlType) {
                            # Reduce requested rights with honoured rights
                            $mask = $mask -band -bnot $_.FileSystemRights
                            if ($mask -eq 0) {
                                # All requested rights honoured
                                return
                            }
                        }
                        else {
                            # Inverse accessControlType
                            if (($mask -band $_.FileSystemRights) -ne 0) {
                                # One or more requested rights explicitly not honoured
                                return
                            }
                        }
                    }

                    if ($mask -eq 0) { $ir.Value }
                }
            }
        )
    }

    return $ht
}


function ModifyFileSecurityDescriptor {
    param (
        [hashtable] $SystemParams,
        [hashtable] $FunctionParams,
        [string]    $FullName
    )

    $folder = Get-Item -LiteralPath $FullName

    $sd = $folder.GetAccessControl()
    $is_sd_modified = $false

    if ($FunctionParams.ContainsKey('InheritanceEnable')) {
        $sd.SetAccessRuleProtection(-not $FunctionParams.InheritanceEnable, $false)    # !preserveInheritance
        $is_sd_modified = $true
    }

    if ($FunctionParams.ContainsKey('Owner')) {
        $sd.SetOwner((New-Object $SystemParams.principal_type $FunctionParams.Owner))
        $is_sd_modified = $true
    }

    $access_profiles = GetAccessProfiles $SystemParams $FunctionParams

    $irs_of_aps = GetIdentityReferencesMatchingAccessProfiles $sd $access_profiles $SystemParams.principal_type

    $irs_involved = @(
        foreach ($key in $irs_of_aps.Keys) { $irs_of_aps[$key] }
    )

    foreach ($ap in $access_profiles) {
        $irs_involved += $FunctionParams[$ap.property_name]
    }

    $irs_involved = @( $irs_involved | Sort-Object -Unique )

    for ($ix = 0; $ix -lt $irs_involved.Count; $ix++) {
        $irs_involved[$ix] = New-Object $SystemParams.principal_type $irs_involved[$ix]
    }

    $is_acl_purged = $false

    foreach ($ap in $access_profiles) {
        if (! $is_acl_purged) {
            $sd.GetAccessRules($true, $false, $SystemParams.principal_type) | ForEach-Object {
                if ($irs_involved.Contains($_.IdentityReference)) {
                    $sd.RemoveAccessRuleAll($_)
                    $is_sd_modified = $true
                }
            }

            $is_acl_purged  = $true
        }

        foreach ($ir in $FunctionParams[$ap.property_name]) {
            $ace = New-Object System.Security.AccessControl.FileSystemAccessRule((New-Object $SystemParams.principal_type $ir), $ap.fileSystemRights, $ap.inheritanceFlags, $ap.propagationFlags, $ap.accessControlType)
            $sd.AddAccessRule($ace)
            $is_sd_modified = $true
        }
    }

    if ($is_sd_modified) {
        LogIO info "SetAccessControl" -In -FullName $FullName #-AclObject $sd
            #$rv = Set-Acl -PassThru -LiteralPath $FullName -AclObject $sd
            #$job = Start-Job -ScriptBlock { Set-Acl -LiteralPath $args[0] -AclObject $args[1] } -ArgumentList @($FullName, $sd)
            $rv = $folder.SetAccessControl($sd)
        LogIO info "SetAccessControl" -Out #$rv

        $rv
    }
}



function GetItemsWithDepth {
    param (
        [string]$Path,
        [int]$Depth,
        [array]$Excludes
    )

    # Helper function to recursively list items
    function InternalGetItems {
        param (
            [string]$CurrentPath,
            [int]$CurrentDepth,
            [array]$Exclude
        )
		
        if ($CurrentDepth -ge 0) {
            # Attempt to list the current directory's contents
            $test = $CurrentPath
            try {
                Log debug "Reading $($CurrentPath)"
                try { $items = Get-ChildItem -LiteralPath $CurrentPath -Directory -Force -ErrorAction Stop | ForEach-Object { 
                    foreach ($exclude in $Excludes) {
                        if ($_.FullName -ilike $exclude) { return }
                    }
                    $_
                } } catch { 
                    $error = "Failed to access contents of: [$($CurrentPath)] - $_"
                    Log error $error
					throw $error
                } 

                # Output the items from the current directory
                $items

                # If the depth allows, recurse into subdirectories
                if ($CurrentDepth -gt 0) {
                    foreach ($item in $items) {
                        if ($item.PSIsContainer) {
                            InternalGetItems -CurrentPath $item.FullName -CurrentDepth ($CurrentDepth - 1) -Exclude $Exclude
                        }
                    }
                }
            } catch {
                throw $_
            }
        }
    }

    # Start the recursive listing from the initial path and depth
    InternalGetItems -CurrentPath $Path -CurrentDepth $Depth
}