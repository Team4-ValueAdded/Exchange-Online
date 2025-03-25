# Connect to Exchange Online
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

"" > .\log.txt # Clear out the previous log, or create one if it doesn't exist.

# Path to the CSV file containing email addresses
Write-Host "Enter the filename (or Path) of the CSV containing users to add: " -ForegroundColor Cyan -NoNewline
$CsvFilePath = Read-Host

# Specify the distribution group to add users to
Write-Host "Specify the distribution group to add users to in Display Name format: " -ForegroundColor Cyan -NoNewline
$DistributionGroupName = Read-Host

try {
    Write-Host "Importing List from '$CsvFilePath'..." -ForegroundColor Yellow
    # Import the list of email addresses from the CSV file
    $UsersToAdd = Import-Csv -Path $CsvFilePath
    Write-Host $UsersToAdd
} catch {
    Write-Host $($_.Exception.Message) -ForegroundColor Red
    exit
}

# Determine if the EmailAddress header exists
$headers = ($UsersToAdd | Get-Member -MemberType NoteProperty).Name  
if("EmailAddress" -in $headers){  
   
} else {
    Write-Host "First cell in CSV must be 'EmailAddress'" -ForegroundColor Red
    exit
}

Write-Host "Adding users to distribution group '$DistributionGroupName'..." -ForegroundColor Yellow

# Loop through each email address in the CSV file
foreach ($User in $UsersToAdd) {
    $UserToAdd = $User.EmailAddress

    Write-Host "====================================================================" -ForegroundColor Yellow
    Write-Host "Processing $UserToAdd..." -ForegroundColor Yellow

    # Check if the user is already a member of the distribution group
    $Members = Get-DistributionGroupMember -Identity $DistributionGroupName
    if ($Members | Where-Object {$_.PrimarySmtpAddress -eq $UserToAdd}) {
        Write-Host "$UserToAdd is already a member of $DistributionGroupName" -ForegroundColor Green
        "$UserToAdd is already a member of $DistributionGroupName" >> .\log.txt
    } else {
        # Add the user to the distribution group
        try {
            Add-DistributionGroupMember -Identity $DistributionGroupName -Member $UserToAdd -ErrorAction Stop
            Write-Host "Added $UserToAdd to $DistributionGroupName" -ForegroundColor Cyan
            "Added $UserToAdd to $DistributionGroupName" >> .\log.txt
        } catch {
            # Check to see if the account exists
            try {
                $unused = Get-Recipient -Identity $UserToAdd -ErrorAction Stop # This variable is for only checking to see if the email exists. It is discarded.
                Write-Host "Failed to add $UserToAdd to $DistributionGroupName. $($_.Exception.Message)" -ForegroundColor Red
                "Failed to add $UserToAdd to $DistributionGroupName. $($_.Exception.Message)" >> .\log.txt

            # If not, output so
            } catch {
                Write-Host "Failed to add $UserToAdd to $DistributionGroupName. The account doesn't exist." -ForegroundColor Red
                "Failed to add $UserToAdd to $DistributionGroupName. The account doesn't exist." >> .\log.txt
            
            }
        }
    }
}
$localdir = Get-Location
Write-Host "Done. Review the log at $localdir\log.txt" -ForegroundColor Cyan
