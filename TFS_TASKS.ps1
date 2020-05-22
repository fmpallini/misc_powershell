<#
  v1.0

  USAGE:
  - Open this script on PowerShell ISE;
  - Remember to allow the execution of scripts in PowerShell before trying to run it. Eg. 'Set-ExecutionPolicy Unrestricted -Scope CurrentUser'
  - Configure your TFS base URL, it may need your collection and project ids/paths.
  - Use the commands bellow;
  
  #############################
  #Example generic APIs calls #
  #############################

  #Params in order (space-separated) - operation (string), resource (string), body (string)
  callTFS "GET" "workitems/632895"
  callTFS "POST" "wiql" '{"query": "SELECT * FROM workitems WHERE [System.Id] = 594447"}'
  callTFS "GET" "workitemtypes"

  Official doc of the TFS Rest API -> https://docs.microsoft.com/en-us/rest/api/azure/devops/?view=azure-devops-rest-4.1

  ########################
  #Example task creation #
  ########################

  #Create task function
  #Params in order (space-separated) - parentId (long), title (string), activity (string), remaining hours (long), description (string)
  createTask 1234 "A&D" "Development" 1 ""

#>

#TFS Enviroment
$base_url = "http://xxx/_apis/wit/"
$api_version = "4.1"

#Save credentials
$credentials_file = "tfs_credentials.txt"

if(![System.IO.File]::Exists("$HOME\$credentials_file")){
   $cred = Get-Credential -Message "Enter your username (with domain) and password.`r`nExample User: domain\user"
   if(!$cred)
   {
     Exit
   }

   $cred = @{
    user = $cred.UserName
    pass = $cred.Password | ConvertFrom-SecureString
   }
   
   $cred | ConvertTo-Json | Set-Content -Path "$HOME\$credentials_file"
}
else
{
    $cred = Get-Content -Path "$HOME\$credentials_file" -Raw | ConvertFrom-Json
}

$cred.pass = $cred.pass | ConvertTo-SecureString

#control variables
$global:lastParentId = 0
$global:areaPath = ""
$global:iterationPath = ""

#Functions
function callTFS()
{
   Param
   (
         [Parameter(Mandatory=$true, Position=0)]
         [ValidateSet("POST","GET","PATCH","DELETE","PUT")]
         [string]$method,
         [Parameter(Mandatory=$true, Position=1)]
         [string]$resource,
         [Parameter(Mandatory=$false, Position=2)]
         [string]$body
   )

   $URL = $base_url + $resource + (&{If(($resource.ToCharArray()) -contains '?') {"&"} Else {"?"}}) + "api-version=" + $api_version
   $URI = New-Object System.Uri($URL,$true)
   
   $ptrPass = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred.pass)
   $Creds = new-object System.Net.NetworkCredential($cred.user, [Runtime.InteropServices.Marshal]::PtrToStringAuto($ptrPass))
   [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptrPass)

   $request = [System.Net.HttpWebRequest]::Create($URI)
   $request.Credentials = $Creds
   $request.Headers.Add("cache-control","no-cache")
   $request.ContentType = "application/json-patch+json"
   $request.Accept = "application/json"
   $request.Method = $method

   if($body)
   {
      $requestBody = [byte[]][char[]]$body
      $request.ContentLength = $requestBody.Length 
      
      $stream = $request.GetRequestStream()
      $stream.Write($requestBody, 0, $requestBody.Length)
   }

   [System.Net.HttpWebResponse] $response = [System.Net.HttpWebResponse] $request.GetResponse()
   $reader = [IO.StreamReader] $response.GetResponseStream()
   $output = $reader.ReadToEnd()
   $response.Close()

   return $output
}

function createTask()
{
   Param
   (
         [Parameter(Mandatory=$true, Position=0)]
         [long]$parentId,
         [Parameter(Mandatory=$true, Position=1)]
         [string]$title,
         [Parameter(Mandatory=$true, Position=2)]
         [ValidateSet("Analysis","Automation","Deployment","Development","Documentation","Environment Setup","Review","Testing","Training/Knowledge Transfer")]
         [string]$activity,
         [Parameter(Mandatory=$false, Position=3)]
         [ValidateRange(0,40)]
         [int]$remaining,
         [Parameter(Mandatory=$false, Position=4)]
         [string]$description
   )

   if($global:lastParentId -ne $parentId)
   {
      $parent = callTFS "GET" "workitems/$parentId" | ConvertFrom-Json
      $global:areaPath = $parent.fields.'System.AreaPath' -replace '\\','\\'
      $global:iterationPath = $parent.fields.'System.IterationPath' -replace '\\','\\'
      $global:lastParentId = $parent.id
   }

   $body =
   '[
       {
          "op": "add",
          "path": "/fields/System.Title",
          "from": null,
          "value": ' + (ConvertTo-Json($title)) + '
       },
       {
          "op": "add",
          "path": "/fields/System.AreaPath",
          "from": null,
          "value": "' + $global:areaPath + '"
       },
       {
          "op": "add",
          "path": "/fields/System.IterationPath",
          "from": null,
          "value": "' + $global:iterationPath + '"
       }
    '

    if($activity)
    {
      $body += ',
      {
         "op": "add",
         "path": "/fields/Microsoft.VSTS.Common.Activity",
         "from": null,
         "value": "' + $activity + '"
      }'
    }

    if($remaining)
    {
      $body += ',
      {
         "op": "add",
         "path": "/fields/Microsoft.VSTS.Scheduling.RemainingWork",
         "from": null,
         "value": ' + $remaining + '
      }'
    }

    if($description)
    {
      $body += ',
      {
         "op": "add",
         "path": "/fields/System.Description",
         "from": null,
         "value": ' + (ConvertTo-Json($description)) + '
      }'
    }

    $body +=',
      {
         "op": "add",
         "path": "/relations/-",
         "value": {
               "rel": "System.LinkTypes.Hierarchy-Reverse",
               "url": "' + $base_url + 'workItems/' + $parentId + '"
          }
      }
      ]'

      return (callTFS "POST" "workitems/`$Task" $body | ConvertFrom-Json).id
}
