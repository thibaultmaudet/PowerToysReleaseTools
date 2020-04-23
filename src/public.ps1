function Get-PowerToysReleaseCurrent {
    begin {
        Write-Verbose "[$((Get-Date).TimeofDay) BEGIN  ] Starting: $($MyInvocation.Mycommand)"
    }

    process {
        $data = Get-Data @PSBoundParameters

        if ($PSVersionTable.ContainsKey("GitCommitID")) {
            $local = $PSVersionTable.GitCommitID
        }
        else {
            $Local = $PSVersionTable.PSVersion
        }

        if ($data.tag_name) {
            [PSCustomObject]@{
                Name = $data.name
                Version = $data.tag_name
                Released = $($data.published_at -as [datetime])
                LocalVersion = $local
            }
        }
    }

    end {
        Write-Verbose "[$((Get-Date).TimeofDay) END    ] Ending: $($MyInvocation.Mycommand)"
    }
}

Function Get-PowerToysReleaseSummary {
    [cmdletbinding()]
    [OutputType([System.String[]])]
    param(
        [Parameter(HelpMessage = "Display as a markdown document")]
        [switch]$AsMarkdown
    )

    begin {
        Write-Verbose "[$((Get-Date).TimeofDay) BEGIN  ] Starting: $($MyInvocation.Mycommand)"
    }

    process {
        $data = Get-Data

        $dl = $data.assets |
        Select-Object @{Name = "Filename"; Expression = {$_.name}},
        @{Name = "Updated"; Expression = {$_.updated_at -as [datetime]}},
        @{Name = "SizeMB"; Expression = {$_.size / 1MB -as [int]}}

        if ($AsMarkdown) {
            $tbl = (($DL | ConvertTo-Csv -notypeInformation -delimiter "|").Replace('"', '') -Replace '^', "|") -replace "$", "|`n"

            $out = @"
# $($data.Name.trim())
$($data.body.trim())
## Downloads
$($tbl[0])|---|---|---|
$($tbl[1..$($tbl.count)])
Published: $($data.Published_At -as [datetime])
"@
        }
        else {
            $out = @"
-----------------------------------------------------------
$($data.Name)
Published: $($data.Published_At -as [datetime])
-----------------------------------------------------------
$($data.body)
-------------
| Downloads |
-------------
$($DL | Out-String)"
"@
        }

        $out
    }

    end {
        Write-Verbose "[$((Get-Date).TimeofDay) END    ] Ending: $($MyInvocation.Mycommand)"
    }
}

Function Get-PowerToysReleaseAsset {
    begin {
        Write-Verbose "[$((Get-Date).TimeofDay) BEGIN  ] Starting: $($MyInvocation.Mycommand)"
    }

    process {
        try {       
            Write-Verbose "[$((Get-Date).TimeofDay) PROCESS] Getting normal assets"
            $data = Get-Data -ErrorAction stop            
            
            [regex]$rx = "(?<file>[p|P]ower[t|T]oys[-|_]\d.*)\s+-\s+(?<hash>\w+)"

            $r = $rx.Matches($data.body)
            $r | ForEach-Object -Begin {
                $h = @{}
            } -Process {
                $f = $_.groups["file"].value.trim()
                $v = $_.groups["hash"].value.trim()

                if (-not ($h.ContainsKey($f))) {
                    Write-Verbose "[$((Get-Date).TimeofDay) PROCESS] Adding $f [$v]"
                    $h.add($f, $v )
                }
                else {
                    Write-Verbose "[$((Get-Date).TimeofDay) PROCESS] Ignoring duplicate asset: $f [$v]"
                }
            }

            Write-Verbose "[$((Get-Date).TimeofDay) PROCESS] Found $($data.assets.count) downloads"

            $assets = $data.assets |
            Select-Object @{Name = "FileName"; Expression = {$_.Name}},
            @{Name = "Format"; Expression = {
                    $_.name.split(".")[-1]
                }
            },
            @{Name = "SizeMB"; Expression = {$_.size / 1MB -as [int32]}},
            @{Name = "Hash"; Expression = {$h.item($_.name)}},
            @{Name = "Created"; Expression = {$_.Created_at -as [datetime]}},
            @{Name = "Updated"; Expression = {$_.Updated_at -as [datetime]}},
            @{Name = "URL"; Expression = {$_.browser_download_Url}},
            @{Name = "DownloadCount"; Expression = {$_.download_count}}

            if ($assets.filename) {
                $assets
            }
            else {
                Write-Warning "Failed to find any release assets using the specified critiera."
            }
        }
        catch {
            throw $_
        }
    }

    end {
        Write-Verbose "[$((Get-Date).TimeofDay) END    ] Ending: $($MyInvocation.Mycommand)"
    }
}

Function Save-PowerToysReleaseAsset {
    [cmdletbinding(DefaultParameterSetName = "All", SupportsShouldProcess)]
    [OutputType([System.IO.FileInfo])]
    param(
        [Parameter(Position = 0, HelpMessage = "Where do you want to save the files?")]
        [ValidateScript( {
            if (Test-Path $_) {
                $True
            }
            else {
                Throw "Cannot validate path $_"
            }
        })]
        [string]$Path = ".",
        [Parameter(ParameterSetName = "All")]
        [switch]$All,

        [Parameter(ParameterSetName = "file", ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [object]$Asset,

        [switch]$Passthru
    )

    begin {
        Write-Verbose "[$((Get-Date).TimeofDay) BEGIN  ] Starting: $($MyInvocation.Mycommand)"
    }

    process {
        Write-Verbose "[$((Get-Date).TimeofDay) PROCESS] Using Parameter set $($PSCmdlet.ParameterSetName)"

        if ($PSCmdlet.ParameterSetName -match "All|Family") {
            Write-Verbose "[$((Get-Date).TimeofDay) PROCESS] Getting latest releases from $uri"
            Try {
                $data = Get-PowerToysReleaseAsset -ErrorAction Stop
            }
            Catch {
                Write-Warning $_.exception.message
                #bail out
                Return
            }
        }

        Switch ($PSCmdlet.ParameterSetName) {
            "All" {
                Write-Verbose "[$((Get-Date).TimeofDay) PROCESS] Downloading all releases to $Path"
                foreach ($asset in $data) {
                    Write-Verbose "[$((Get-Date).TimeofDay) PROCESS] ...$($Asset.filename) [$($asset.hash)]"
                    $target = Join-Path -Path $path -ChildPath $asset.filename
                    Download-File -source $asset.url -Destination $Target -hash $asset.hash -passthru:$passthru
                }
            }
            "File" {
                Write-Verbose "[$((Get-Date).TimeofDay) PROCESS] ...$($asset.filename) [$($asset.hash)]"
                $target = Join-Path -Path $path -ChildPath $asset.fileName
                Download-File -source $asset.url -Destination $Target -hash $asset.hash -passthru:$passthru
            }
        }
    }

    End {
        Write-Verbose "[$((Get-Date).TimeofDay) END    ] Ending: $($MyInvocation.Mycommand)"
    }
}

function Install-PowerToys {
    [cmdletbinding(SupportsShouldProcess)]
    Param(
        [Parameter(HelpMessage = "Specify the path to the download folder")]
        [string]$Path = $env:TEMP,
        [Parameter(HelpMessage = "Specify what kind of installation you want. The default if a full interactive install.")]
        [ValidateSet("Full", "Quiet", "Passive")]
        [string]$Mode = "Full"
    )
    Begin {
        Write-Verbose "[$((Get-Date).TimeofDay) BEGIN  ] Starting $($myinvocation.mycommand)"
    }

    Process {
        if (($psedition -eq 'Desktop') -OR ($PSVersionTable.platform -eq 'Win32NT')) {
            if ($PSBoundParameters.ContainsKey("WhatIf")) {
                Write-Verbose "[$((Get-Date).TimeofDay) PROCESS] Creating a dummy file for WhatIf purposes"
                $filename = Join-Path -path $Path -ChildPath "whatif-preview.msi"
            }
            else {
                Write-Verbose "[$((Get-Date).TimeofDay) PROCESS] Saving download to $Path "
                $install = Get-PowerToysReleaseAsset | Save-PowerToysReleaseAsset -Path $Path -Passthru
                $filename = $install.fullname
            }

            Write-Verbose "[$((Get-Date).TimeOfDay) PROCESS] Using $filename"

            $inParams = @{
                Path = $filename
                Mode = $Mode
                ErrorAction = "stop"
            }

            if ($pscmdlet.ShouldProcess($filename, "Install PowerToys using $mode mode")) {
                Install-MSI @inParams
            }
        }
        else {
            Write-Warning "This will only work on Windows platforms."
        }
    }

    end {
        Write-Verbose "[$((Get-Date).TimeofDay) END    ] Ending $($myinvocation.mycommand)"
    }

}