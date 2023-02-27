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
    
}
