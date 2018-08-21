$Modules = @("Visio")

foreach ($Module in $Modules) {
    if (-not (Get-Module -ListAvailable -Name $Module)) { 
        if (-not(Get-Module -ListAvailable -Name PowerShellGet)) { 
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force 
        }
        Install-Module $Module  -Force 
    }
    Import-Module $Module
}

New-VisioApplication
New-VisioDocument
$basic_u = Open-VisioDocument -Filename "basic_u.vss"
$master = Get-VisioMaster -Name "Rectangle" -Document $basic_u

$shape = New-VisioShape -Masters $master -Points 4,5

Set-VisioText -Text "Hello World" -Shape $shape