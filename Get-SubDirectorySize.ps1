[CmdletBinding()]
Param
(
        [Parameter(ValueFromPipeline=$True)]            
        [Alias('directory','dir','d')]
        [String[]]$DirectoryNames = [String](Get-Location)
)

Begin
{
    $output = @()
}

Process
{
    ForEach($directory in $DirectoryNames)
    {
        <# Removes the Microsoft.PowerShell.Core\FileSystem if Get-Location is called on a network drive #>
        $directorySplit = ($directory.Split("::"))
        If($directorySplit.length -ge 2)
        {
            $directory = $directorySplit[2]
        }

        $colItems = Get-ChildItem $directory | Where-Object {$_.PSIsContainer -eq $true} | Sort-Object
        foreach ($i in $colItems)
        {
            $subFolderItems = Get-ChildItem $i.FullName -Recurse -Force | Where-Object {$_.PSIsContainer -eq $false} | Measure-Object -Property Length -Sum
            
            $temp = New-Object PSObject
            $temp | Add-Member -MemberType NoteProperty -Name Folder -Value $i.Name
            $temp | Add-Member -MemberType NoteProperty -Name Root -Value $directory
            $temp | Add-Member -MemberType NoteProperty -Name "Size(MB)" -Value ("{0:N2}" -f ($subFolderItems.Sum / 1MB))

            $output += $temp
        }
    }
}
End
{
    Return $output
}
