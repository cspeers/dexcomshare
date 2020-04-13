
$Script:Dexcom_Settings=[PSCustomObject]@{
    ApplicationId = "d89443d2-327c-4a6f-89e5-496bbb0317db"
    ShareApi = @{
        EU="shareous1.dexcom.com";
        US="share1.dexcom.com"
    }
}

#region Types

enum DexcomTrend
{
    None
    DoubleUp
    SingleUp
    FortyFiveUp
    Flat
    FortyFiveDown
    SingleDown
    DoubleDown
    NotComputable
    OutOfRange
}

class DexcomShareGlucoseEntry
{
    [DateTime]$DT
    [DateTime]$ST
    [DexcomTrend]$Trend
    [int]$Value
    [DateTime]$WT
}

class NightScoutGlucoseEntry
{
    [int]$sgv
    [long]$date
    [string]$dateString
    [int]$trend
    [string]$direction
    [string]$device='share2'
    [string]$type='sgv'
}

#endregion

<#
    .SYNOPSIS
        Convert a trend reading to a direction
#>
function ConvertTo-Direction
{
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline=$true,Mandatory=$true)][DexcomTrend[]]$Trend
    )
    process
    {
        foreach ($item in $Trend)
        {
            switch ($item)
            {
                0{Write-Output "None"}
                1{Write-Output "DoubleUp"}
                2{Write-Output "SingleUp"}
                3{Write-Output "FortyFiveUp"}
                4{Write-Output "Flat"}
                5{Write-Output "FortyFiveDown"}
                6{Write-Output "SingleDown"}
                7{Write-Output "DoublDown"}
                8{Write-Output "NotComputable"}
                9{Write-Output "OutOfRange"}
            }
        }
    }
}

<#
    .SYNOPSIS
        Converts a dexcom glucose entry to a NightScout entry
#>
function ConvertTo-Nightscout
{
    [OutputType([NightScoutGlucoseEntry[]])]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][DexcomShareGlucoseEntry[]]$GlucoseEntries
    )
    begin
    {
        $epoch = New-Object System.DateTime(1970, 1, 1, 0, 0, 0, 0);
    }
    process
    {
        foreach ($GlucoseEntry in $GlucoseEntries)
        {
            $ns=[NightScoutGlucoseEntry]::new()
            $ns.dateString=$GlucoseEntry.WT.ToString('s')
            $ns.date=[Math]::Floor(($GlucoseEntry.WT - $epoch).TotalSeconds)
            $ns.trend=$GlucoseEntry.Trend
            $ns.sgv=$GlucoseEntry.Value
            $ns.direction=$GlucoseEntry.Trend.ToString()
            Write-Output $ns
        }
    }
}

<#
    .SYNOPSIS
        Retrieves a Dexcom Share Session Id
    .PARAMETER Credential
        The Dexcom Share Account
    .PARAMETER AuthUri
        The Authentication Endpoint
    .PARAMETER Account
        Dexcom Share Account
    .PARAMETER AccountName
        Dexcom Share Account Name
    .PARAMETER AccountSecret
        Dexcom Share Account Password
#>
function Get-DexcomShareSessionId
{
    [CmdletBinding(DefaultParameterSetName='Credential')]
    param
    (
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$ApplicationId=$Script:Dexcom_Settings.ApplicationId,
        [Parameter(ValueFromPipelineByPropertyName=$true,Mandatory=$true,ParameterSetName='Clear')]
        [Alias('UserName')]
        [string]$AccountName,
        [Parameter(ValueFromPipelineByPropertyName=$true,Mandatory=$true,ParameterSetName='Clear')]
        [Alias('Password')]
        [string]$AccountSecret,
        [Parameter(ValueFromPipelineByPropertyName=$true,Mandatory=$true,ParameterSetName='Credential')]
        [Alias('DexcomAccount')]
        [pscredential]$Account,    
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [uri]$AuthUri=$Script:Dexcom_Settings.ShareApi.US
    )
    
    begin
    {
        $RequestBldr=New-Object UriBuilder "https://${AuthUri}"
        $RequestBldr.Path='/ShareWebServices/Services/General/LoginPublisherAccountByName'
        $SessionId=[Guid]::Empty
        if($PSCmdlet.ParameterSetName -eq 'Credential') {
            $nc=$Account.GetNetworkCredential()
            $AccountName=$nc.UserName
            $AccountSecret=$nc.Password
        }
    }
    
    process 
    {
        try
        {
            $RequestParams=@{
                Headers=@{
                    'User-Agent'="Dexcom Share/3.0.2.11 PowerShell/$($PSVersionTable.PSVersion)"
                    'Accept'='application/json'
                }
                ContentType='application/json'
                Uri=$RequestBldr.Uri
                UseBasicParsing=$true
                Method='POST'
            }
            $Response=Invoke-WebRequest @RequestParams -Body $(@{applicationId=$ApplicationId;accountName=$AccountName;password=$AccountSecret}|ConvertTo-Json)
            $SessionId=$Response.Content.Replace('"',"")
        }
        catch {
            $ErrorContent=$_.Exception.Message
            #Should I unwind the exception?
            if($PSVersionTable.PSVersion.Major -lt 6) {
                $ErrorStream = $itemException.Response.GetResponseStream()
                $ErrorStream.Position = 0
                $StreamReader = New-Object System.IO.StreamReader($ErrorStream)
                try
                {
                    $ErrorContent = $StreamReader.ReadToEnd()
                    $StreamReader.Close()
                }
                catch
                {

                }
                finally
                {
                    $StreamReader.Close()
                }
            }
            else {
                $ErrorContent=$_.ErrorDetails.Message
            }
            Write-Error "An Error Occurred!`t${ErrorContent}"
        }
        if($SessionId -eq [Guid]::Empty) {
            throw 'No Session Id was returned! Please check your account details.'
        }
        Write-Output $SessionId
    }

    end {}
}

<#
    .SYNOPSIS
        Retrieves the set of glucose values for the desired period
    .PARAMETER ApiEndpoint
        The Dexcom share Api Endpoint
    .PARAMETER SessionId
        The Dexcom Share Session Id
    .PARAMETER AsNightScout
        Whether to return the records in nightscout format
    .PARAMETER IntervalInMinutes
        The interval of records to retrieve
    .PARAMETER MaxCount
        The maximum amount of records to retrieve
#>
function Get-DexcomShareLatestGlucoseValues
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [Uri]$ApiEndpoint=$Script:Dexcom_Settings.ShareApi.US,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
        [string[]]$SessionId,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [Alias('Interval')]
        [int]$IntervalInMinutes=1440,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)]
        [int]$MaxCount=[Math]::Floor($IntervalInMinutes/12),
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [switch]$AsNightscout
    )
    begin
    {
        $Results=@()
        $RequestBuilder=New-Object UriBuilder "https://${ApiEndpoint}"
        $RequestBuilder.Path='/ShareWebServices/Services/Publisher/ReadPublisherLatestGlucoseValues'
        $RequestParams=@{
            ContentType='application/json'
            Headers=@{'User-Agent'="Dexcom Share/3.0.2.11 PowerShell/$($PSVersionTable.PSVersion)";'Accept'='application/json'}
            Method='Post'
            UseBasicParsing=$true
        }
    }
    process
    {
        foreach ($item in $SessionId)
        {
            try
            {
                #?sessionID=e59c836f-5aeb-4b95-afa2-39cf2769fede&minutes=1440&maxCount=1"
                $RequestBuilder.Query="sessionID=${item}&minutes=${IntervalInMinutes}&maxCount=${MaxCount}"
                $RequestParams.Uri=$RequestBuilder.Uri
                $Response=Invoke-WebRequest @RequestParams
                [DexcomShareGlucoseEntry[]]$RawVals=$Response.Content|ConvertFrom-Json
                if($AsNightscout) {
                    $Results+=$RawVals|ConvertTo-Nightscout
                }
                else
                {
                    $Results+=$RawVals
                }
            }
            catch
            {
                #Should I catch this and keep going, or bother unwinding?
                $ErrorContent=$_.Exception.Message
                #Should I unwind the exception?
                if($PSVersionTable.PSVersion.Major -lt 6) {
                    $ErrorStream = $itemException.Response.GetResponseStream()
                    $ErrorStream.Position = 0
                    $StreamReader = New-Object System.IO.StreamReader($ErrorStream)
                    try
                    {
                        $ErrorContent = $StreamReader.ReadToEnd()
                        $StreamReader.Close()
                    }
                    catch
                    {
    
                    }
                    finally
                    {
                        $StreamReader.Close()
                    }
                }
                else {
                    $ErrorContent=$_.ErrorDetails.Message
                }
                Write-Warning "An Error Occurred!`t${ErrorContent}"
            }
        }
    }
    end
    {
        Write-Output $Results
    }
}

New-Alias -Name 'Login-DexcomShare' -Value 'Get-DexcomShareSessionId'
New-Alias -Name 'Get-LatestBSG' -Value 'Get-DexcomShareLatestGlucoseValues'