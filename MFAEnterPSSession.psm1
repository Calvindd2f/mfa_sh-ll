function Invoke-MFAEnterPSSession {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [PSCredential]$Credential
    )

    # Step 1: MFA logic. 
    # You can use Azure AD, Google Authenticator, Authy, or another MFA provider.
    # Custom function `Request-MFAOTP` that sends an OTP to the user and returns the expected OTP value.
    $expectedOTP = Request-MFAOTP

    # Step 2: Prompt the user to enter the OTP received
    $userInput = Read-Host "Enter the OTP received"

    # Step 3: Verify the OTP
    if ($userInput -eq $expectedOTP) {
        Write-Host "OTP verification successful. Proceeding with Enter-PSSession..."
        
        # Step 4: Establish the remote session using Enter-PSSession
        if ($Credential) {
            Enter-PSSession -ComputerName $ComputerName -Credential $Credential
        } else {
            Enter-PSSession -ComputerName $ComputerName
        }
    } else {
        Write-Host "Invalid OTP. Access denied."
    }
}

function Request-MFAOTP ([string]$UserEmail) {
    # Replace this URL with the URL of your RequestMFAOTP Azure Function
    if(!($UserEmail)){
        $functionUrl = 'https://your-function-app.azurewebsites.net/api/otp?code=your-function-key'
    } else{
        $functionUrl = 'https://your-function-app.azurewebsites.net/api/otp?code=your-function-key&email=' + $UserEmail
    }

    try {
        $response = Invoke-WebRequest -Uri $functionUrl -Method Get
        $result = $response | ConvertFrom-Json

        if ($response.StatusCode -eq 200) {
            return $result.otp
        } else {
            throw "Error calling Azure Function: $($result.message)"
        }
    } catch {
        Write-Error $_.Exception.Message
        return $null
    }
}
