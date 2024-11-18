Write-Host "Reading Email Files...."

#Define Outlook object and namespace:
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

#Retrieves default email and inbox in Outlook:
$inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)
$items = $inbox.Items


#Creates an array object to store parsed data from for loop:
$reportEmails = New-Object System.Collections.ArrayList




#Parse the inbox for details:
Write-Host "Getting emails, please do no close...."
foreach ($item in $items){

		#if iterator is of type mail Item:
		if($item -is [Microsoft.Office.Interop.Outlook.MailItem]){
			
			#create null variable as placeholder:
			$senderEmail=$null	
			
			#Try to retrieve senders email address: 
			try{
				#Try to get the senders email directly:
				if($item.Sender -ne $null){
				 $sender = $item.Sender 
				 $senderEmail = $item.SenderEmailAddress


		
					#Checks if AddressEntry is not null:
					if(-not $senderEmail -and $sender.AddressEntry -ne $null){
						
						#assign variable to AddressEntity:
						$senderEmail = $sender.AdderssEntry.Address
						
					}
					
					#Use mail service provider to locate source of maile:
					if(-not $senderEmail -and $sender.AddressEntry -ne $null){
						$exchangeUser = $sender.AddressEntry.GetExchangeUser()


						if($exchangeUser -ne $null){
							$senderEmail = $exchangeUser.PrimaryAddress
						}
					}
				}	
			}catch{
				$senderEmail = "Unkonwn Sender"
				Write-Host "Caught unknown sender"
				Write-Host "Sender: $($item.Sender)"
				Write-Host "Email Address: $($item.SenderEmailAddress)"
								
			}

					$null=$reportEmails.Add([PSCustomObject]@{
						From = $senderEmail
						Subject = $item.Subject
						Recieved = $item.ReceivedTime
					})	
		}			
					
}	
		


$desktopPath  = [Environment]::GetFolderPath('Desktop')

$outputFile = Join-Path $desktopPath "SourceStats.csv"

$reportEmails | Export-Csv -Path $outputFile -NoTypeInformation -Encoding  UTF8

Write-Host "Results saved to $outputFile"

Write-Host "Closing..."

#Clean up Outlook COM object to prevent issues:
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
Remove-Variable -Name outlook -Force
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
