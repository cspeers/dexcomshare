# dexcomshare PowerShell Module
## A simple wrapper for Dexcom Share Requests

- Get-DexcomShareSessionId
    + (ParameterSet Credential) Account 
    ```powershell
    $Account=New-Object PSCredential 'yourdexcomaccount',(ConvertTo-SecureString -String 'yourpassword' -AsPlainTextForce)
    $SessionId=Get-DexcomshareSessionId -Account $Account
    ```                              
    + (ParameterSet Clear) AccountName
    + (ParameterSet Clear) AccountSecret
    ```powershell
    $SessionId=Get-DexcomshareSessionId -AccountName 'yourdexcomaccount' -AccountSecret 'yourpassword'
    ```    
- Get-DexcomShareLatestGlucoseValues
    + SessionId
    + IntervalInMinutes
    + MaxCount
    + AsNightScout
    ```powershell
    Get-DexcomShareLatestGlucoseValues -SessionId $SessionId -IntervalInMinutes 1440 -MaxCount 120 -AsNightScout
    ```