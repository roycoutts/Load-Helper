$result = Invoke-RestMethod -Uri "https://coronavirus-19-api.herokuapp.com/countries"
$result | Out-GridView
