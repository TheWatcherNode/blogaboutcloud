#'@',

# https://docs.microsoft.com/en-gb/office365/enterprise/prepare-for-directory-synchronization?redirectSourcePath=%252fen-us%252farticle%252fPrepare-to-provision-users-through-directory-synchronization-to-Office-365-01920974-9e6f-4331-a370-13aea4e82b3e
cls
 $array = @('~', '!', '#', '$', '%', '^', '&', '(', ')', '-', '.+', '=', '}', '{', '\', '/', '|', ';', ',', ':', '<', '>', '"')
 $samaccountarray = @('[', '\', '"', '|' , ',' , '/', ':', '<', '>', '+', '=', ';', ']')
 foreach ($char in $array) {
 Write-Host "Please Wait... Detecting",$char," in samaccountname" -ForegroundColor Yellow
 $objects = Get-distributiongroup
 foreach ($object in $Objects)
 {
 try {
  if ($object.SamAccountName -like "*$char*") 
 {
 Write-Host "Special Character",$char,"detected in SamAccountName",$object.samaccountname -ForegroundColor Red
 
 }
 else
 {
 #Write-Host "Special Character",$char," not detected in " $object.UserPrincipalName
 }
 }
 catch
 {
 Write-Host "Great News!! we was unable to detect",$char,"in samaccountnames for all Distribution List" -ForegroundColor Green
 }
 }
 }
 