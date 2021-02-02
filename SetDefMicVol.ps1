Import-Module AudioDeviceCmdlets;
$CSVFilePath = "MicScript.csv"
Function Set-VolumeAfterCheck
{
    [CmdletBinding()]
    Param
    (
        [int]$SelectedVolume
    )
    While($True)
    {
        if((Get-AudioDevice -RecordingVolume) -ne $SelectedVolume)
        {
            Set-AudioDevice -RecordingVolume $SelectedVolume;
        }
        Start-Sleep -Seconds 5;
    }
}
if ((Test-Path -Path "$CSVFilePath") -eq $False)
{
    New-Item -Path "$CSVFilePath";
    Add-Content -Path "$CSVFilePath" -Value "Microphone Volume";
    Add-Content -Path  "$CSVFilePath" -Value 60;
    $SelectedVolume = (Import-CSV -Path "$CSVFilePath")."Microphone Volume"
    Set-VolumeAfterCheck -SelectedVolume $SelectedVolume;
}
else
{
    $SelectedVolume = (Import-CSV -Path "$CSVFilePath")."Microphone Volume"
    Set-VolumeAfterCheck -SelectedVolume $SelectedVolume;
}