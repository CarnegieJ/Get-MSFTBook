<#
.SYNOPSIS
   Retrieves Microsoft eBook from web location.
.DESCRIPTION
   Retrieves Microsoft eBook from web location and outputs information about the retrieval.
.PARAMETER MsftBook
   The URL of the Microsoft eBook to retrieve.
.PARAMETER DestinationFolder
   The optional filepath location to save the retrieved eBook.
.PARAMETER NoDownload
   If this switch is specified the eBook will not be retrieved.
.INPUTS
   System.String
   You can pipe a string that contains the URL of the Microsoft eBook to retrieve.
.OUTPUTS
   System.Management.Automation.PSCustomObject
   The cmdlet returns a PS custom object with the following member properties:
     Destination  - Full filepath name to retrieved eBook
     File         - Web server response file name
     Item         - Web request URL
     Source       - Web server response URI
.NOTES
   Name:    Get-MSFTBook
   Author:  Carnegie Johnson, MCT
   Email:   CarnegieJ@msn.com
   Twitter: @CarnegieJ
   Created: 2016-Jul-12
   Copyright (c) 2016 Carnegie Johnson

   Thank you to Microsoft Director of Sales Excellence - Eric Ligman
   for the inspiration to create this script. God bless you and our Microsoft family.
.LINK
   https://IAYFConsulting.com/Carnegie
.EXAMPLE
   PS C:\>Get-MSFTBook -MsftBook 'http://ligman.me/1omCrM6' 

   Description
   -----------
   This command will retrieve the Microsoft eBook "PowerShell_Examples_v4.pdf" to the temporary folder,
   the cmdlet returns the custom object.
.EXAMPLE
   PS C:\>Get-MSFTBook -MsftBook 'http://ligman.me/1omCrM6' -DestinationFolder 'C:\eBooks'

   Description
   -----------
   This command will retrieve the Microsoft eBook "PowerShell_Examples_v4.pdf" to the "C:\eBooks" folder
   and create the folder if it does not exist, the cmdlet returns the custom object.
.EXAMPLE
   PS C:\>Get-MSFTBook -MsftBook 'http://ligman.me/1omCrM6' -NoDownload

   Description
   -----------
   This command uses the NoDownload switch and will NOT retrieve the Microsoft eBook "PowerShell_Examples_v4.pdf",
   the cmdlet returns the custom object.
#>
Function Get-MSFTBook {
  [CmdletBinding(DefaultParameterSetName='Pset1' 
                , PositionalBinding=$true)]
  Param (
    # Web resource(s) to download
    [Parameter(Mandatory=$true 
            , Position=0
            , HelpMessage="Please specify URL."
            , ParameterSetName='Pset1')]
    [ValidateNotNullOrEmpty()]
    [string[]]$MsftBook, 
    # Folder path to store download
    [Parameter(Mandatory=$false,
              Position=1,
              ParameterSetName='Pset1')]
    [string]$DestinationFolder = $env:TEMP,
    [Parameter(Mandatory=$false,
              Position=2,
              ParameterSetName='Pset1')]
    [switch]$NoDownload)
  Begin {
    $www = New-Object System.Net.Webclient
    $output = [PSCustomObject]@{
      Item = [string]$MsftBook;
      File = $null;
      Source = $null;
      Destination = $null}

    $resource = ''
    $nl = [System.Environment]::NewLine
  }

  Process {
    Trap [System.Net.WebException] {
      $errX = $_.Exception
      $errmsg = $nl + $errX.ToString() + $nl
      $errmsg += ("  Argument= {0}" -f $MsftBook) + $nl
      $status = [System.Net.WebExceptionStatus]$errX.Status
      
      switch ($true) {
        $($status -match 'ProtocolError') {
          $errResponse = [System.Net.WebResponse]$errX.Response
          $errHttpResponse = [System.Net.HttpWebResponse]$errX.Response
          $errmsg += "ProtocolError Status Code = {0}{1}" -f $errResponse.StatusCode.Value__, $nl
          $errmsg += "Description: {0} <{1}>" -f $errResponse.StatusDescription, $errHttpResponse.ResponseUri
          $errmsg += "Message: {0}" -f $weberr.Message; break}
        default {
          $errmsg += ("  Status = {0}" -f $status.ToString())
        }
      }
      $errmsg
      Break
    } # End Trap
    If ($MsftBook -match '(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)') {
      Try {
        If (-not (Test-Path -Path $DestinationFolder)) {
          New-Item -ItemType Directory -Path $DestinationFolder -Force
        }
        $requestURI = New-Object System.Uri($MsftBook)
        $request = [System.Net.WebRequest]::CreateHttp($requestURI)
        $response = $request.GetResponse()
        If ($response -eq $null) {
          return
        }
        $responseURI = $response.ResponseUri
        $resource = (($responseURI -split "/")[-1]).Split("?", 2)[0]
        $response.Dispose()
        $dest = Join-Path -Path $DestinationFolder -ChildPath $resource
        If(-not $NoDownload) {
          $www.DownloadFile($responseURI, $dest)
        }
        $output.File = [string]$resource
        $output.Source = [string]$responseURI
        $output.Destination = [string]$dest
      } # End Try
      Catch {
        $errX = $_.Exception
        $errmsg = $errX.ToString() + $nl + ("  Argument= {0}" -f $MsftBook)
        $errmsg
      }
    } # Endif URL
  }

  End {
    $www.Dispose()
    $output
  }
}

# Source of "Download All" list for Microsoft FREE eBooks
$MsftBooks = "http://www.mssmallbiz.com/ericligman/Key_Shorts/MSFTFreeEbooks.txt"
# Get contents of the book list and split each line into an object collection
$URLs = (Invoke-WebRequest -URI $MsftBooks).Content.Split("`n").Trim()
# Filter object colletion for URL's in HTTP:// or FTP:// protocol formats
$URLs = $URLs | ?{$_ -match '(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)'}
# Filter duplicate URL's
$URLs = $URLs | Sort-Object | Get-Unique
# Retrieve each eBook into "Downloads" folder and output to grid
$URLs | %{Get-MSFTBook -MsftBook $_ -DestinationFolder "C:\Downloads\Microsoft\eBooks"} | Out-GridView
