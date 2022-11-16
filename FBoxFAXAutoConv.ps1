#
# FBoxFAXAutoConv.ps1
# Gert Michael (Kruemelino) 2022
# based on asynchonous FileSystemWatcher Example: https://powershell.one/tricks/filesystem/filesystemwatcher
# https://github.com/Kruemelino/FBoxFAXAutoConv
#

# Pure Powershell cmdlets to read and write dBase .DBF-files. https://github.com/Delapro/PSDBF
# get file from github repository and store it in the same folder as this script
."$PSScriptRoot\DBFReadWrite.ps1"

# specify the path to the folder you want to monitor:
$FritzFaxPath = "$($env:APPDATA)\FRITZ!\Fax\"

# specify which files you want to monitor
$FileFilter = "*.sff"  

# specify whether you want to monitor subfolders as well:
$IncludeSubfolders = $false

# specify the file or folder properties you want to monitor:
$AttributeFilter = [IO.NotifyFilters]::FileName, [IO.NotifyFilters]::LastWrite 

# specify the folder for the file transfer
$DestinationFolder = "E:\tmp\FBoxFAXAutoConv\"

# FAXTools SSF2TIFF <c> Shamrock 2016
$SSF2TIFFPath = "E:\tmp\FBoxFAXAutoConv\SFF2TIFF.EXE"

# FAXTools TIFF2PDF <c> Shamrock 2016
$TIFF2PDFPath = "E:\tmp\FBoxFAXAutoConv\TIFF2PDF.EXE"

# Returns a new filename. Informations of the fax were taken from the FritzFax.Dbf
function GetFileName($sffPath) 
{
    $NewFileName = ""
    
    # Write-Host "sff: $sffPath" -ForegroundColor Green

    # Write-Host "dbf:" (Join-Path -Path $FritzFaxPath -ChildPath 'FritzFax.Dbf') -ForegroundColor Green

    $a=Use-DBF (Join-Path -Path $FritzFaxPath -ChildPath 'FritzFax.Dbf')

    $daten = foreach($nr in $a.ListAll()) {$a.Goto($nr); new-object -property $a.Fields -typename psobject}

    $a.Close()

    # Write-Host ($daten | foreach {($_.DATEI.Trim())}) -ForegroundColor Red

    $req = $daten | where {$_."DATEI".Trim() -eq $sffPath}


    # check if a entry for that sff file exists
    If ($req) {
        # create ne filename.
        $Datum = $req.psobject.properties["DATUM"].value.Trim()
        $Zeit = $req.psobject.properties["ZEIT"].value.Trim()
        $TelNr = $req.psobject.properties["RUFNUMMER"].value.Trim()
        
        $NewFileName = "$Datum`_$Zeit`_$TelNr.sff" 
    } Else {
        Write-Host "Error: No entry for $sffPath in the database." -ForegroundColor Red
        $NewFileName = "$(get-date -Format yyyyMMdd)`_" + (Get-Item $sffPath).Basename + ".sff"
    }
    
    return $NewFileName
}

try
{
    $watcher = New-Object -TypeName System.IO.FileSystemWatcher -Property @{
               Path = $FritzFaxPath
               Filter = $FileFilter
               IncludeSubdirectories = $IncludeSubfolders
               NotifyFilter = $AttributeFilter }

        # define the code that should execute when a change occurs:
        $action = 
        {   
            # change type information:
            $details = $event.SourceEventArgs
    
            Write-Host ""
            Write-Host ("{0} was {1} at {2}" -f $details.FullPath, $details.ChangeType, $event.TimeGenerated) -ForegroundColor DarkYellow
    
            # you can also execute code based on change type here:
            switch ($details.ChangeType)
            {
                'Changed'  { "CHANGE" }
                'Created'  { 

                    # Rename and Copy the sff-file to $DestinationFolder and return the new file
                    # $faxfile = Copy-Item -Path $details.FullPath -Destination $DestinationFolder -Force -PassThru
                    $faxfile = Copy-Item -Path $details.FullPath -Destination (Join-Path -Path $DestinationFolder -ChildPath (GetFileName($details.FullPath))) -Force -PassThru
                    Write-Host "Convert sff-file to tif-image: $faxfile"
                    
                    # Convert the ssf-file to a tif-image using the FAXTools SSF2TIFF <c> Shamrock 2016
                    & $SSF2TIFFPath $faxfile.FullName

                    # assign the created tif-image to the $faxfile object
                    $faxfile = Get-ChildItem (Join-Path -Path $DestinationFolder -ChildPath ($faxfile.Basename + ".tif"))
                    
                    Write-Host "Convert tif-image to pdf-document: $faxfile"

                    # Convert the ssf-file to a tif-image using the FAXTools TIFF2PDF <c> Shamrock 2016
                    & $TIFF2PDFPath $faxfile.FullName

                    # delete the previous created tif-image
                    Remove-Item $faxfile
                    
                    # assign the created pdf document to the $faxfile object
                    $faxfile = Get-ChildItem (Join-Path -Path $DestinationFolder -ChildPath ($faxfile.Basename + ".pdf"))
                    
                    Write-Host "Created pdf-document: $faxfile" -ForegroundColor Green
                }
                'Deleted'  { "DELETED"}
                'Renamed'  { "RENAMED"}
        
                # any unhandled change types surface here:
                default   { Write-Host $_ -ForegroundColor Red -BackgroundColor White }
            }
        }

        # subscribe your event handler to all event types that are
        # important to you. Do this as a scriptblock so all returned
        # event handlers can be easily stored in $handlers:
        $handlers = . {
            # Register-ObjectEvent -InputObject $watcher -EventName Changed  -Action $action 
            Register-ObjectEvent -InputObject $watcher -EventName Created  -Action $action 
            # Register-ObjectEvent -InputObject $watcher -EventName Deleted  -Action $action 
            # Register-ObjectEvent -InputObject $watcher -EventName Renamed  -Action $action 
        }

        # monitoring starts now:
        $watcher.EnableRaisingEvents = $true

        Write-Host "Watching for changes to $FritzFaxPath"

        # since the FileSystemWatcher is no longer blocking PowerShell
        # we need a way to pause PowerShell while being responsive to
        # incoming events. Use an endless loop to keep PowerShell busy:
        do
        {
            # Wait-Event waits for a second and stays responsive to events
            # Start-Sleep in contrast would NOT work and ignore incoming events
            Wait-Event -Timeout 1

            # write a dot to indicate we are still monitoring:
            Write-Host "." -NoNewline
        
        } while ($true)
    }
finally
    {
        # this gets executed when user presses CTRL+C:
  
        # stop monitoring
        $watcher.EnableRaisingEvents = $false
  
        # remove the event handlers
        $handlers | ForEach-Object {
            Unregister-Event -SourceIdentifier $_.Name
        }
  
        # event handlers are technically implemented as a special kind
        # of background job, so remove the jobs now:
        $handlers | Remove-Job
  
        # properly dispose the FileSystemWatcher:
        $watcher.Dispose()

        # properly dispose the PrintDocument:
        $PrintDocument.Dispose()

        Write-Warning "Event Handler disabled, monitoring ends."
    }