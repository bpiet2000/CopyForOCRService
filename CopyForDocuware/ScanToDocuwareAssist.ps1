#
#Due to the fact that DocuWare deletes files upon import this script is going to make a copy of a file and send it to the Docuware Import Folder.
#The origional will be moved into a folder to indicate that it has been scaned. 
#
#Script will only import PDF documents
#
#****If you need to change any of the copy or read file paths please edit DocuWareAssistSettings.csv****
#For safety make sure the file path uses Fully Qualified Domain names of thier share locations


function send-NJETemail ([system.string]$emailBody) {
	
#We are importing the encrypted email user namme and password 
#IMPORTANT:  Windows Encrypts this information on a per computer basis.  A new O365Smtp.xml file has to be generated
#for each computer this script runs on
	$emailCred = Import-Clixml (Join-Path -Path $PSScriptRoot -ChildPath "O365Smtp.xml")
	
	
	$emailOptions = @{
	From = $emailCred.UserName
#If you want to change who the emails go to edit this line below.  
#Each e-mail address needs to be enclosed in quotes.  
	To = "IT Group <itsupport@n-jet.com>"
	Subject = "Scan to DocuWare Assistant"
	Body = $emailBody
	SMTPServer = "smtp.office365.com"
	Port = "587"
	Credential = $emailCred
	}
	Send-MailMessage @emailOptions -UseSsl

}



[system.diagnostics.EventLog]::CreateEventSource(“DocuWare PowerShell Script”, “Application”)
$DocumentTypesToImport = import-csv -Path (Join-Path -Path $PSScriptRoot -ChildPath "DocuWareAssistSettings.csv")
if ($DocumentTypesToImport -eq $null) {
	#We exit the script if the variable is empty
	$errorMessage = "There is an error in DocuWareAssistSettings.csv.  `n Does it exist?  `n Does it have any entries?"
	send-NJETemail $errorMessage
	exit
}

# Loop through every document type listed in DocuWareAssitSettings.csv
# As long as the three file paths are valid you can list as many as you want. 
foreach ($DocType in $DocumentTypesToImport) {
	
	$AllLocationsValid = $true
	if ((Test-Path -Path $DocType.FileToImportPath) -eq $false) {
		$errorMessage = "Error while processings $DocType. `n Can't locate Path of Files to Import."
		$AllLocationsValid = $false
		send-NJETemail $errorMessage
	}
	if ((Test-Path -Path $DocType.DocuwareImportPath) -eq $false) {
		$errorMessage = "Error while processings $DocType. `n Can't locate DocuWare import path."
		$AllLocationsValid = $false
		send-NJETemail $errorMessage
	}
	if ((Test-Path -Path $DocType.DocuwareImportPath) -eq $false) {
		$errorMessage = "Error while processings $DocType. `n Can't locate where to store documents after import."
		$AllLocationsValid = $false
		send-NJETemail $errorMessage
	}
	if ($AllLocationsValid) {
		#Get a list of items that have a *.pdf extension.
		#Files get copied to the DocuWareImportPath
		#Then the origional files get moved to the PathToMoveFileAfterImport
		$fileList = Get-ChildItem -Path $DocType.FileToImportPath -Filter "*.pdf"
		foreach ($file in $fileList) {
			Copy-Item $file.FullName -Destination $DocType.DocuwareImportPath
			Move-Item $file.FullName -Destination $DocType.PathToMoveFileAfterImport
		}
		$DocTypeCompleted = $DocType.DocumentType
		$logMessage = "A Powershell Script has copied $($DocType.DocumentType) files for import into DocuWare."
		Write-EventLog -LogName Application -Source “DocuWare PowerShell Script” -EventId 45233 -EntryType Information -Message $logMessage
	}
	
}

# SIG # Begin signature block
# MIIPVwYJKoZIhvcNAQcCoIIPSDCCD0QCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUpApJAgc7jnu5s9dVWasW8ZL7
# 0++gggmpMIIEmTCCA4GgAwIBAgIPFojwOSVeY45pFDkH5jMLMA0GCSqGSIb3DQEB
# BQUAMIGVMQswCQYDVQQGEwJVUzELMAkGA1UECBMCVVQxFzAVBgNVBAcTDlNhbHQg
# TGFrZSBDaXR5MR4wHAYDVQQKExVUaGUgVVNFUlRSVVNUIE5ldHdvcmsxITAfBgNV
# BAsTGGh0dHA6Ly93d3cudXNlcnRydXN0LmNvbTEdMBsGA1UEAxMUVVROLVVTRVJG
# aXJzdC1PYmplY3QwHhcNMTUxMjMxMDAwMDAwWhcNMTkwNzA5MTg0MDM2WjCBhDEL
# MAkGA1UEBhMCR0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQMA4GA1UE
# BxMHU2FsZm9yZDEaMBgGA1UEChMRQ09NT0RPIENBIExpbWl0ZWQxKjAoBgNVBAMT
# IUNPTU9ETyBTSEEtMSBUaW1lIFN0YW1waW5nIFNpZ25lcjCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAOnpPd/XNwjJHjiyUlNCbSLxscQGBGue/YJ0UEN9
# xqC7H075AnEmse9D2IOMSPznD5d6muuc3qajDjscRBh1jnilF2n+SRik4rtcTv6O
# KlR6UPDV9syR55l51955lNeWM/4Og74iv2MWLKPdKBuvPavql9LxvwQQ5z1IRf0f
# aGXBf1mZacAiMQxibqdcZQEhsGPEIhgn7ub80gA9Ry6ouIZWXQTcExclbhzfRA8V
# zbfbpVd2Qm8AaIKZ0uPB3vCLlFdM7AiQIiHOIiuYDELmQpOUmJPv/QbZP7xbm1Q8
# ILHuatZHesWrgOkwmt7xpD9VTQoJNIp1KdJprZcPUL/4ygkCAwEAAaOB9DCB8TAf
# BgNVHSMEGDAWgBTa7WR0FJwUPKvdmam9WyhNizzJ2DAdBgNVHQ4EFgQUjmstM2v0
# M6eTsxOapeAK9xI1aogwDgYDVR0PAQH/BAQDAgbAMAwGA1UdEwEB/wQCMAAwFgYD
# VR0lAQH/BAwwCgYIKwYBBQUHAwgwQgYDVR0fBDswOTA3oDWgM4YxaHR0cDovL2Ny
# bC51c2VydHJ1c3QuY29tL1VUTi1VU0VSRmlyc3QtT2JqZWN0LmNybDA1BggrBgEF
# BQcBAQQpMCcwJQYIKwYBBQUHMAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20w
# DQYJKoZIhvcNAQEFBQADggEBALozJEBAjHzbWJ+zYJiy9cAx/usfblD2CuDk5oGt
# Joei3/2z2vRz8wD7KRuJGxU+22tSkyvErDmB1zxnV5o5NuAoCJrjOU+biQl/e8Vh
# f1mJMiUKaq4aPvCiJ6i2w7iH9xYESEE9XNjsn00gMQTZZaHtzWkHUxY93TYCCojr
# QOUGMAu4Fkvc77xVCf/GPhIudrPczkLv+XZX4bcKBUCYWJpdcRaTcYxlgepv84n3
# +3OttOe/2Y5vqgtPJfO44dXddZhogfiqwNGAwsTEOYnB9smebNd0+dmX+E/CmgrN
# Xo/4GengpZ/E8JIh5i15Jcki+cPwOoRXrToW9GOUEB1d0MYwggUIMIIC8KADAgEC
# AhBPK+ajHGOUhEHXdlrvF8b3MA0GCSqGSIb3DQEBBQUAMBwxGjAYBgNVBAMMEU5K
# RVQgQ29kZSBTaWduaW5nMB4XDTE4MDYwNDE0Mzg0NloXDTE5MDYwNDE0NTg0Nlow
# HDEaMBgGA1UEAwwRTkpFVCBDb2RlIFNpZ25pbmcwggIiMA0GCSqGSIb3DQEBAQUA
# A4ICDwAwggIKAoICAQC8m17vrX1nzDAFEtgSMux3e7ihp4tRYa2MbTqh+rjHI8IA
# pj76s1eKuEBvYIjUkUWMUssAQLD8l+o5rqVJb33zQCr2jzRRa4NfjtV0LFKJ/V6N
# 0JIKvFYpZlg/FH684eqVbx0+GUv/IWzHi3lrTG+zSKIHnuF5sq3Cd5h2dS5BxSSn
# Dh1ZoP4+O61hE2Kyqe2c9XaVKToDWP3sj0wACxghLLXJlK8JOXZyIa2ZDJX59lCG
# wtfWOsePSjMyWmeQ3aMWpDVQOQuMWUyIORucdsUypZcgCYT0TiQMK80uXwbUb025
# rskyvIKNhJQk2wcCu/7X14Q1EdaOIGa8YAgKCLJKQ9BiIUI1bof0qthgMkgC9Rev
# ZXui3HmAIpdAe0xmoShG1VWgU9ZplZgx83tv+/YLDqd/8RRdw5z0+VjCXnkIM3f+
# DUmnJZ3o87SzlJDpRClfao87kJRDnHK/yEXKfXEruRlLlgxM0Ks+i33y+EaM+D66
# 0w2P02dge/MRDcXn2q2jGuNnPtgL7fxdCjG1opnz4Zpf1XTDdGLUJDO9QEhfOlHu
# L9r7bsM/AAmLO1ELGJunD/gAIiQZq+ZdJ1DnrR7DD8eRw9kQkD4A73Pij3HD++H3
# AUN4E2YClshMqOAEyrSE7wdEohs4nVJKS4c8QpAyY9nCAcVRVmHPCTeWGDnFfQID
# AQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwHQYD
# VR0OBBYEFH4R0dAe8qSCKaRvXo6Rh34NmpNVMA0GCSqGSIb3DQEBBQUAA4ICAQCP
# 3iKR4Hv4p7yVYKkV7Sx0rsZhPd+X4I2Mkdb3hBJQA8RUnnksnt9JMgtogtU/HoXj
# Y0EDwjAkHvBTXQv5UdDsDJ4bA7Q8CWA2fPkgU9IYJb5KSU4ee+rN8IUVwQkhLRrG
# wDCpoSA54R5wwc48Zy6C3mpGXcszceBLGPyNMVG0qwHbw4Ow41NPb1V7SSjjclDC
# fgsYBlyhnPiCHYZwlSccsSGgk+WGlVprmVw/KJFHRobkvx1BOX2+yxYvnjRvSasF
# 2i4S3NCU/7XsR2hXU4j6H+lXH5dUoslevsVGjPHEvAWFWVJnC7tn/i4601kv4dNV
# Lly2czbD6TKiY7FwzvjCmlILNB+pkW0Olt5/1+NU1d6yYD38mRH3hD9M9OPCU+PJ
# 3GKG/VOYTCFTsSxGRDH1uX2cXQQo/nokiAAic0THbObfBcYKCsBZMp409FrliJuQ
# YtSfLHZb/ztPXTp6kVqUAbqkD8xSXo777O1rk6oIMM/NfD9wOysbXpawvBI7RlSi
# Hg/Ko20MPGo8NB9HNUIjoONfyMyEec7ts6xjcz49M6n/+njKlfmFtdLvs/PAS7s/
# +qC7D6FlpA0gAUJTfHrysd2yxZPX1mje2J+IrpOyRo3AH7MrTdDyQWSfUWDjtR57
# l7PwY+Ap5avRBP8fe6HyKCuR5qGcVlSzGQRc6nmNLDGCBRgwggUUAgEBMDAwHDEa
# MBgGA1UEAwwRTkpFVCBDb2RlIFNpZ25pbmcCEE8r5qMcY5SEQdd2Wu8XxvcwCQYF
# Kw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkD
# MQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJ
# KoZIhvcNAQkEMRYEFADnHmBInYK6u0/ZkmhDrAHZGmyqMA0GCSqGSIb3DQEBAQUA
# BIICALEcvWU7gKPSN9IgUji8KiB3o4SwqGBAtk7kP4/LlGlBn4u+mt6cgoUYk/58
# ocBdJu3MErjWZnEdtRVKE9d0rhHnmT1EIAPz+CRwph5EuskBqvOrGh+szkoYikbb
# JTPTRl/FUKTJqV+Ccbi44iZeb+rgGPMUogQRU10fRM+KfwEqxfXrzWcjrMRVTiD1
# xNExDWismMjVADDG0F5+2h1v4jX15dNPnBKsanAMzLXcHaR2+onCkiQrv32bgLLM
# Sob/FkK3l+x3XVpIXwk9A0dlphZHuh1U269ABC1Vsq2jD+0s6D4PKHOjqJyNbZjA
# Pps0ZBNzNB0kOZUSwCMR15iOjxCKlIp3MG46387RRggq9V6GYo8tjHS5zipIKf6f
# vu7Svl53cLSwlBHDwqArJjLujuge3ARUWo4UCb/+y4F9+k/6n1Xn1K41TiGnVo3O
# Z6H691qmVoTitPSiBINGCY/Z0ZnJcFAl83yTsLIp8lf/gFojtZtS/Fo8oj2NBrEB
# ZXmLsBgQDDtuQB/6cUe0Pyx95PjDLKxJALBzVl6OxjDPkqrxc5K9YlrpBKWXCey+
# xZVbjeZHY01mBRll2eoRc8wj0dDV96XNoqmxoMdADdhLFEPGSbzZlugTsfuxqZp5
# lNzgyXJLHXmiJ9U580hmq5Wrwz7BpsoTvIL9SkzjYXmrn9lqoYICQzCCAj8GCSqG
# SIb3DQEJBjGCAjAwggIsAgEBMIGpMIGVMQswCQYDVQQGEwJVUzELMAkGA1UECBMC
# VVQxFzAVBgNVBAcTDlNhbHQgTGFrZSBDaXR5MR4wHAYDVQQKExVUaGUgVVNFUlRS
# VVNUIE5ldHdvcmsxITAfBgNVBAsTGGh0dHA6Ly93d3cudXNlcnRydXN0LmNvbTEd
# MBsGA1UEAxMUVVROLVVTRVJGaXJzdC1PYmplY3QCDxaI8DklXmOOaRQ5B+YzCzAJ
# BgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0B
# CQUxDxcNMTgwNzE3MDA1MDU5WjAjBgkqhkiG9w0BCQQxFgQUVMEhKgfifMvvii0s
# 6ARe3bVpjXkwDQYJKoZIhvcNAQEBBQAEggEAip4CcC1hL0jUfTi2ejBuYy6ecry2
# 1IA8ak0DBg1+0LsigXPbweiQe3N763krX4Ge80TNY/MzqaeF3dr9kE/cb+3mDiqR
# VTCdQBOyAT6kfGMOqu4plvpy/H5ww9reGstl7lfJMQRtcirMd1SuL+F/2YE9X7Zx
# Pg3mOOVRK3kuZpMk6Ev3nIIjwFE4mYF9gS+N0u4LA5T47QwtqACQunJ/HpcaI8Oq
# T5/03uqXd2vRNdslc/147wPdSmA+o7LQ9AeHqewWhgeqPGQHhTCjW4VQ94weiqWQ
# PMtdPtFktOb1qMB/9WaZgE8TaE3hEaM/Dz8MW6OF4xEMqb7Q9R/DM49kOw==
# SIG # End signature block
