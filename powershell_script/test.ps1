$postParams = @{client_id='50fc1ee4af7c3cfbac7dd9b56f952fbb';client_secret='9b50cc37736d4576b64ece1bc7e78c55';grant_type='password';username='vesselit@oocl.com';password='Sea##856'}
$result = Invoke-WebRequest -Headers $headers -Uri https://api.ic.peplink.com/api/oauth2/token -Method POST -Body $postParams
$json = $result.Content| ConvertFrom-Json
$uri = "https://api.ic.peplink.com/rest/o/gYr0QD/g/3/d?device_count=0&has_status=false&access_token="+$json.access_token
$result = Invoke-WebRequest -Uri $uri
$groupDevice = $result.Content | ConvertFrom-Json
$status = @()
foreach ($i in $groupDevice.data){
    $deviceId=$i.id
    $uri = "https://api.ic.peplink.com/rest/o/gYr0QD/g/3/d/"+$deviceId+"/status?access_token="+$json.access_token
    $result = Invoke-WebRequest -Uri $uri
    $deviceStatus = $result.Content | ConvertFrom-Json
    $params = New-Object -TypeName PSObject
    $params = $deviceStatus.data
    $params | Add-Member -MemberType NoteProperty -Name server_ref -Value $deviceStatus.server_ref 
    $params | Add-Member -MemberType NoteProperty -Name device_id -Value $i.id
    $params | Add-Member -MemberType NoteProperty -Name device_name -Value $i.name
    $status += $params
}
if(!(Test-Path  "E:\VesselCloud\pepwave\current_$(Get-Date -Format ddMMyyyy).csv")){
    $status | convertto-csv -notype | out-file -Append -Encoding utf8 -filepath E:\VesselCloud\pepwave\current_$(Get-Date -Format ddMMyyyy).csv
}
else{
    $status | convertto-csv -notype | Select -Skip 1 | out-file -Append -Encoding utf8 -filepath E:\VesselCloud\pepwave\current_$(Get-Date -Format ddMMyyyy).csv
}
