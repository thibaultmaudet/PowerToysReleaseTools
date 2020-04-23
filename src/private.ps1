function Get-Data {
    [cmdletbinding()]
    Param(
        [switch]$Preview
    )

    $uri = "https://api.github.com/repos/microsoft/PowerToys/releases"

    Write-Verbose "[$((Get-Date).TimeofDay) $($myinvocation.mycommand)] Getting current release information from $uri"
    $get = Invoke-Restmethod -uri $uri -Method Get -ErrorAction stop

    if ($Preview) {
        Write-Verbose "[$((Get-Date).TimeofDay) $($myinvocation.mycommand)] Getting latest preview"
        ($get).where( {$_.prerelease}) | Select-Object -first 1
    }
    else {
        Write-Verbose "[$((Get-Date).TimeofDay) $($myinvocation.mycommand)] Getting latest stable release"
        ($get).where( { -NOT $_.prerelease}) | Select-Object -first 1
    }
}

function Download-File {
    [cmdletbinding(SupportsShouldProcess)]
    Param([string]$Source, [string]$Destination, [string]$hash, [switch]$Passthru)

    Write-Verbose "[$((Get-Date).TimeofDay) $($myinvocation.mycommand)] $Source to $Destination"

    if ($pscmdlet.ShouldProcess($Destination, "Downloading $source")) {
        Invoke-Webrequest -Uri $source -UseBasicParsing -DisableKeepAlive -OutFile $Destination
        Write-Verbose "[$((Get-Date).TimeofDay) $($myinvocation.mycommand)] Comparing file hash to $hash"
        $f = Get-FileHash -Path $Destination -Algorithm SHA256
        
        if ($f.hash -ne $hash) {
            Write-Warning "Hash mismatch. $Destination may be incomplete."
        }

        if ($passthru) {
            Get-Item $Destination
        }
    }
}

function Install-Msi {
    [cmdletbinding(SupportsShouldProcess)]
    Param(
        [Parameter(Mandatory, HelpMessage = "The full path to the MSI file")]
        [ValidateScript({Test-Path $_})]
        [string]$Path,
        [Parameter(HelpMessage = "Specify what kind of installation you want. The default if a full interactive install.")]
        [ValidateSet("Full", "Quiet", "Passive")]
        [string]$Mode = "Full",
        [Parameter(HelpMessage = "Enable PowerShell Remoting over WSMan.")]
        [switch]$EnableRemoting,
        [Parameter(HelpMessage = "Enable the PowerShell context menu in Windows Explorer.")]
        [switch]$EnableContextMenu
    )

    Write-Verbose "[$((Get-Date).TimeofDay) $($myinvocation.mycommand)] Creating Start-Process parameters"

    $modeOption = switch ($Mode) {
        "Full"    {"/qf" }
        "Quiet"   {"/quiet"}
        "Passive" {"/passive"}
    }

    $installOption = "$modeOption REGISTER_MANIFEST=1"

    Write-Verbose "[$((Get-Date).TimeofDay) $($myinvocation.mycommand)] FilePath: $Path"
    Write-Verbose "[$((Get-Date).TimeofDay) $($myinvocation.mycommand)] ArgumentList: $installOption"

    if ($pscmdlet.ShouldProcess("$Path $installOption")) {
        Write-Verbose "[$((Get-Date).TimeofDay) $($myinvocation.mycommand)] Starting installation process"
        Start-Process -FilePath $Path -ArgumentList $installOption
    }

    Write-Verbose "[$((Get-Date).TimeofDay) $($myinvocation.mycommand)] Ending function"
}