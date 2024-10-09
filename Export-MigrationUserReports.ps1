Param
(
[Parameter(Mandatory=$true)]
[System.IO.DirectoryInfo]$RootOutputFolder
)

Function New-SubFolder
{
    Param
    (
    [Parameter(Mandatory=$true)]
    [System.IO.DirectoryInfo]$Name
    )

    $FullPath = Join-Path -Path $RootOutputFolder -ChildPath $Name

    if (-not (Test-Path -Path $FullPath))
    {
        Try
        {
            New-Item -Path $RootOutputFolder -Name $Name -Force -ItemType Directory -Confirm:$false -ErrorAction Stop
        }

        catch
        {
            Write-Host -ForegroundColor Red -Object "Error creating folder $($FullPath). $($_)"
        }
    }
}

$Batches = (Get-MigrationBatch -Status Completed).Identity.Name

Foreach ($Batch in $Batches)
{
    Write-Host -ForegroundColor DarkYellow -Object "Processing Users from Batch $($Batch)`n"

    [System.IO.DirectoryInfo]$FolderPath = New-SubFolder -Name $Batch
    $MigUsers = Get-MigrationUser -BatchId $Batch
    
    Foreach ($MigUser in $MigUsers)
    {
        Write-Host -ForegroundColor Green -Object "Processing User $($MigUser.Identity)"

        [System.IO.FileInfo]$Filename = ($MigUser.Identity.ToString().Replace("@","_") + ".xml")
        $Filepath = Join-Path -Path $FolderPath -ChildPath $Filename
        $Stat = Get-MigrationUserStatistics -Identity $MigUser.Identity -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose" -IncludeSkippedItems
        $Stat | Export-Clixml -Path $Filepath
        $Filename = $null
        $Filepath = $null
        $Stat = $null
    }
}


