function Set-Theme {
    param (
        [Parameter(Mandatory)]
        [ValidateSet("Light", "Dark")]
        [string]$Theme
    )
    if ( $Theme -eq "Dark") {
        $val = 0;
    } else {
        $val = 1;
    }
    Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name AppsUseLightTheme -Value $val
    Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name SystemUsesLightTheme -Value $val

    $explorers = get-explorerpaths;
    Stop-Process -name explorer;
    start-explorers $explorers;

}

function get-explorerpaths () {
    $shell = New-Object -ComObject Shell.Application
    $windows = $shell.Windows()
    $explorerPaths = @()

    foreach ($window in $windows) {
        if ($null -ne $window.Document.Folder.Self.Path) {
            $explorerPaths += $window.Document.Folder.Self.Path
        }
    }

    return $explorerPaths
}
function start-explorers () {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string[]]
        $paths
    )
    foreach ($path in $paths) {
        explorer.exe $path
    }
}
