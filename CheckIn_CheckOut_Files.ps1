Clear-Host

Function Start-Script
{
    Param(
        [Parameter(Mandatory)][String]$SPSite,
        [Parameter(Mandatory)][Boolean]$ReportOnly
    )
    Measure-Command {
        #
        # Variables
        $Script:Report = @()
        #
        # make sure the environment to create logs is available
        $script:HomeDirectory = $env:USERPROFILE+"\Desktop\PowerShell\CheckIn_CheckOut_Files"
        mkdir $HomeDirectory -force | Out-Null
        #
        # Let's get Client Cert that we will use in all the REST Calls!
        $clientCert = Get-ClientCertficiate 
        if($clientCert){
            Get-CheckedOutFiles -cert $clientCert -SPSite $SPSite -ReportOnly $ReportOnly
        }
    }
}
#
# log function
Function log($string, $color) 
{
	    $logfile = "$($HomeDirectory)\CheckIn_CheckOut_$(get-date -format `"yyyyMMdd_hhmmsstt`")_log.txt"
	    if ($null -eq $Color) {$color = "white"}
	    $logDate = Get-Date -Format g
	    $logMessage = "$logDate - " + $string.Exception.Message + "`n" + $string.InvocationInfo.PositionMessage
	    write-host $logMessage -foregroundcolor $color
	    $logMessage| out-file -Filepath $logfile -append
    }
#
# Function get the client clertficate of the user running the script. It currently only grabs the cert for a 
# contractor and a GS Civilian
Function Get-ClientCertficiate 
{
	try{
		#
    	# if we have the client cert, no need to prompt for it again
        $iCerts = Get-ChildItem -Path Cert:\CurrentUser\My
        #
        # loop through all the certs in the store
        foreach($iCert in $iCerts){
            if($iCert.Issuer.Contains("DOD ID CA-51")) { #-or $iCert.Issuer.Contains("DOD EMAIL CA-42")){
                $clientCert = $iCert
            }
        }
		Return $clientCert
	} catch {
		log $Error[0] red
	}
 }
#
#
Function Get-CheckedOutFiles 
{
        Param(
            [Parameter(Mandatory)][X509Certificate]$cert,
            [Parameter(Mandatory)][String]$SPSite,
            [Parameter(Mandatory)][Boolean]$ReportOnly
        )
        try{
            #
            # loop until we have some $Results back. Might need to put a timer
            try{
                # create the web session
                $webSession = Invoke-WebRequest -Method Get -Uri $SPSite -Certificate $cert -SessionVariable session
            } catch {
		        log $Error[0] red 
            }

            #
            # now get all the document libraries in the site
            $SPLibraries = Get-DocumentLibraries $cert $SPSite
            $AllArrays = New-Object System.Collections.ArrayList #@() # this array holds all the arrays
            # make sure $SPLibraries is not empty
            if($null -ne $SPLibraries){
                # create the object to hold a custom build report
                $docLibCount = 1
                Write-Host("There are " + $SPLibraries.Count + " Document Libraries")
                #$HeaderArray = @("This site has " + $SPLibraries.Count + " Document Libraries")
                # loop through all the site libraries on a site
                foreach ($doclib in $SPLibraries){
                    $SingleArray = New-Object System.Collections.ArrayList
                    $SingleArray.Insert(0,$webSession.ParsedHtml.Title)
                    $SingleArray.Insert(1,$SPSite)
                    Write-Host("$($docLibCount)) Working on Document Library: " + $doclib.content.properties.Name)
                    $SingleArray.Insert(2,$docLib.content.properties.Name)
                    #
                    # Call the properties API to get information about the document library. 
                    # This will help us determin if the document library has Check In and Check Out enabled
                    $Results = Invoke-RestMethod -Method Get -Uri ($SPSite + "/_api/" + $doclib.link[4].href) -Certificate $cert -SessionVariable session -ContentType "application/json;odata=verbose"
                    #
                    # Is check in and check out enabled?
                    if($Results.entry.content.properties.vti_x005f_listrequirecheckout -eq "true"){
                        Write-Host("`tThe Document Library, " + $doclib.content.properties.Name + ", has Check-In and Check-Out enabled!") -ForegroundColor Green
                        $SingleArray.Insert(3,$true)
                        #
                        # Does the document library have any documents in it?
                        $ItemCount = $Results.entry.content.properties.vti_x005f_folderitemcount.'#text'
                        if($ItemCount -gt 0){ 
                            #
                            #
                            Write-Host("`tThe Document Library has " + $ItemCount + " items ") -ForegroundColor Green
                            $SingleArray.Insert(4, $ItemCount)
                            #
                            # make the call to the library
                            $doclibURL = ($SPSite + "/_api/web/GetFolderByServerRelativeUrl('" + $doclib.content.properties.Name + "')/Files")
                            $Files = Invoke-RestMethod -Method Get -Uri $doclibURL -Certificate $cert
                            #
                            # Loop through all the files in the document library
                            $i = 0
                            $Files | ForEach-Object{
                                if($i -gt 0){
                                    $SingleArray = New-Object System.Collections.ArrayList
                                    $SingleArray.Insert(0, "")
                                    $SingleArray.Insert(1, "")
                                    $SingleArray.Insert(2, "")
                                    $SingleArray.Insert(3, "")
                                    $SingleArray.Insert(4, "")
                                }else{
                                    $SingleArray.Insert(5, "")
                                    $SingleArray.Insert(6, "")
                                    $SingleArray.Insert(7, "")
                                    $SingleArray.Insert(8, "")
                                    $AllArrays.Add(@($SingleArray)) 
                                    #
                                    # start a new array
                                    $SingleArray = New-Object System.Collections.ArrayList
                                    $SingleArray.Insert(0, "")
                                    $SingleArray.Insert(1, "")
                                    $SingleArray.Insert(2, "")
                                    $SingleArray.Insert(3, "")
                                    $SingleArray.Insert(4, "")
                                }
                                    
                                $i++
                                #
                                # Check to see if the file is checked out. If it is, lets get some info on it and check it back in
                                if($_.content.properties.CheckOutType.'#text' -eq 0){
                                    #
                                    $File = Invoke-RestMethod -Method Get -Uri ($_.Id + "/CheckedOutByUser") -Certificate $cert #-ContentType "application/json;odata=verbose"
                                    $FileURL = $File.entry.id
                                    Write-Host("`tFile, " + $_.content.properties.Name + ", is Checked Out") -ForegroundColor Yellow
                                    Write-Host("`tFile is Checked-Out to: " + $File.entry.content.properties.Title + "; " + $File.entry.content.properties.Email) -ForegroundColor Yellow
                                    $SingleArray.Insert(5, $_.content.properties.Name)
                                    $SingleArray.Insert(6, $true)
                                    $usr = $File.entry.content.properties.Title
                                    $usr = $usr -replace ","
                                    $SingleArray.Insert(7, $usr + " (" + $File.entry.content.properties.Email + ")")
                                    #
                                    # CheckoutType = 0 means it is not checked out
                                    # CheckOutType = 2 means it is checked out online
                                    #
                                    # the following will not execute if $ReportOnly is $true wich is the default setting
                                    #$ReportOnly = $true
                                    if($ReportOnly -ne $true){
                                        # To check-in any document back to SharePoint, A Post is requred. The only way to POST anything to a web page, SharePoint included, 
                                        # is by passing the Request Digest with the commmand. 
                                        # You get the Request Digest by quering the contextinfo property from the site.
                                        Write-Host("`tAttempting To Check the File In! ") -ForegroundColor Yellow
                                        #$SingleArray += (",Attempting To Check the File In! ")
                                        $Response = Invoke-RestMethod -Method Post -uri ($SPSite + "/_api/contextinfo") -Certificate $cert -WebSession $session
                                        #
                                        # Assign the Form Digest so we can send it in the Headers.
                                        $Digest = $Response.getcontextwebinformation.FormDigestValue
                                        $headers = @{
                                            "accept" = "image/gif, image/jpeg, image/pjpeg, application/atom+xml;type=entry, application/x-ms-application, application/xaml+xml, application/x-ms-xbap, */*"; `
                                            "Accept-Language" = "en-US"; `
                                            "X-Requested-With" = "XMLHttpRequest"; `
                                            "IF-MATCH" = "*"; `
                                            "Accept-Encoding" = "gzip, deflate, peerdist";`
                                            "X-RequestDigest" = $Digest; `
                                            "Content-Type" = "application/json;odata=verbose"; `
                                            "Referer" = "https://army.deps.mil"
                                        }
                                        try{
                                        $CheckingIn = Invoke-RestMethod -Method Post `
                                            -Uri ($_.Id + "/CheckIn(" + "comment='2016 Migration', checkintype=0" + ")") `
                                            -Certificate $cert `
                                            -WebSession $session `
                                            -Headers $headers
                                            $SingleArray.Insert(8, "File Checked In Successfully!")
                                        } catch {
                                            $SingleArray.Insert(8, "ERROR: " + $Error[0])
		                                    #Write-Host $Error -ForegroundColor Red
		                                    log $Error[0] Red
                                        }
                                    }else{
                                        Write-Host("`tJust Reporting") -ForegroundColor Yellow
                                        $SingleArray.Insert(8, "Just Reporting")
                                    }
                                }
                                else{
                                    Write-Host("`tFile, " + $_.content.properties.Name + ", is NOT Checked Out") -ForegroundColor Green
                                    $SingleArray.Insert(5, $_.content.properties.Name)
                                    $SingleArray.Insert(6, $false)
                                }
                                $AllArrays.Add(@($SingleArray)) 
                            }
                        }
                        else {
                            Write-Host("The Document Library has 0 items. No further action required") -ForegroundColor Green
                            $SingleArray.Insert(4, "0")
                            $SingleArray.Insert(5, "")
                            $SingleArray.Insert(6, "")
                            $SingleArray.Insert(7, "")
                            $SingleArray.Insert(8, "")
                        }
                    
                    }
                    Write-Host("`tThe Document Library, " + $doclib.content.properties.Name + ", does not have Check-In and Check-Out enabled!") -ForegroundColor Green
                    $SingleArray.Insert(3, $false)
                    $SingleArray.Insert(4, "0")
                    $SingleArray.Insert(5, "")
                    $SingleArray.Insert(6, "")
                    $SingleArray.Insert(7, "")
                    $SingleArray.Insert(8, "")
                    $docLibCount ++
                    $AllArrays.Add(@($SingleArray)) 
                }
            }
            else{
                Write-Host("No Document Libraries Found On This Site") 
            } 
        
        } catch {
            log $Error[0] Red 
        }
        finally {
            $datetime = Get-Date -Format "yyyy_MM_dd_HH_mm_ss";
            $file_name = "CheckOut Files Report " + $datetime + ".csv";
            #$file_path = $HomeDirectory + "\" + $file_name;
            New-Report -arrayObject $AllArrays
        }
    }

#
# Iterates through the site provided and creates a collection of document libraries in the site
Function Get-DocumentLibraries
{
    Param(
        [Parameter(Mandatory)][X509Certificate]$cert,
        [Parameter(Mandatory)][string]$webUrl
    )
    #
    $SPFolders = Invoke-RestMethod -Uri ($SPSite + "/_api/web/Folders") -Method Get -WebSession $session -Certificate $cert
    Return $SPFolders
}

#
#
Function New-Report 
{
    Param(
        [Parameter(Mandatory)][array]$arrayObject
    )
    $Global:daArray = $arrayObject
    #
    #
    try{
        #
        #
        foreach($line in $daArray){
                $newReportItem = New-Object PSObject
                $newReportItem | Add-Member -type NoteProperty -Name "Site Name" -value $line[0]
                $newReportItem | Add-Member -type NoteProperty -Name "Site Url" -value $line[1]
                $newReportItem | Add-Member -type NoteProperty -Name "Document Library" -value $line[2]
                $newReportItem | Add-Member -type NoteProperty -Name "Require CheckOut" -value $line[3]
                $newReportItem | Add-Member -type NoteProperty -Name "Item Count" -value $line[4]
                $newReportItem | Add-Member -type NoteProperty -Name "File" -value $line[5]
                $newReportItem | Add-Member -type NoteProperty -Name "Checked Out" -value $line[6]
                $newReportItem | Add-Member -type NoteProperty -Name "Checked Out By" -value $line[7]
                $newReportItem | Add-Member -type NoteProperty -Name "Action Taken" -value $line[8]
                $Report = [Array]$Report + $newReportItem
        }
        $Report | Export-Csv -Path "C:\temp\ABC_Test.csv" -NoTypeInformation
    } catch {
        log $Error[0] Red
    }
}


Start-Script -SPSite "" -ReportOnly $false
