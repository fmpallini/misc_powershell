<#
  v1.1

  USAGE:
  - Open this script on PowerShell ISE or import into the current Powershell session (eg. ". .\TFS_TASKS.ps1");
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
  
  ########################
  #Example WorkLog       #
  ########################

  #Params in order (space-separated) - issueId (long), additionalWorkHours (float), user (string), date (string), comment (string)
  incrementWorklog 1234 0.5 "Some user" "2020-05-03" "Doing some work"

  ########################################
  #Import WorkLog from TMetrics CSV      #
  ########################################

  #Params in order (space-separated) - csv file path (string)
  importWorklogFromTMetricsCSV ".\file.csv"
  
  #Expected columns:
     Issue Id -> id (with starting # or not)
     Time -> Amound of time in format HH:mm:ss
     User -> identifier
     Day -> Date of that worklog
     Time Entry -> Any comment about the work being done.
#>

#TFS Enviroment
$base_url = "http://xxx/_apis/wit/"
$api_version = "4.1"

#Save credentials
$credentials_file = "tfs_credentials.txt"

if(![System.IO.File]::Exists("$HOME\$credentials_file")){
   $cred = Get-Credential -Message "Enter your username (with domain) and password. The password will be stored on your home folder using SecureString (DPAPI), remember to delete that file everytime you change your password.`r`n`r`nExample User: domain\user"
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
   $request.ContentType = "application/json-patch+json; charset=utf-8"
   $request.Accept = "application/json"
   $request.Method = $method

   if($body)
   {
      $requestBody = [System.Text.Encoding]::UTF8.GetBytes($body);
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

function incrementWorklog()
{
   Param
   (
         [Parameter(Mandatory=$true, Position=0)]
         [long]$issueId,
         [Parameter(Mandatory=$true, Position=1)]
         [ValidateRange(0.25,100)]
         [double]$additionalWorklog,
         [Parameter(Mandatory=$false, Position=2)]
         [string]$user,
         [Parameter(Mandatory=$false, Position=3)]
         [string]$date,
         [Parameter(Mandatory=$false, Position=4)]
         [string]$comment
   )

   #Worklog
   $currentWorklog = (callTFS "GET" "workitems/$issueId" | ConvertFrom-Json).fields.'Microsoft.VSTS.Scheduling.CompletedWork'
   $newWorklog = $currentWorklog + $additionalWorklog

   #Comment
   $note = "Updating Complete Work from: $currentWorklog to: $newWorklog."
   
   if($comment)
   { 
     $note = $note + " - Comment: " + $comment 
   }
   
   if($date)
   { 
     $note = $note + " - on: " + $date 
   }

   if($user)
   {
     $note = $note + " - by " + $user
   }

   $body =
   '[
       {
          "op": "replace",
          "path": "/fields/Microsoft.VSTS.Scheduling.CompletedWork",
          "value": ' + $newWorklog + '
       },
       {
          "op": "add",
          "path": "/fields/System.History",
          "value": ' + (ConvertTo-Json($note)) + '
       }
    ]'

   (callTFS "PATCH" "workitems/$issueId" $body | ConvertFrom-Json).fields.'Microsoft.VSTS.Scheduling.CompletedWork'
}

function importWorklogFromTMetricsCSV()
{
   Param
   (
         [Parameter(Mandatory=$true, Position=0)]
         [string]$filePath
   )

   $csv = Import-Csv $filePath

   Foreach ($i in $csv)
   {
     if($i.'Issue Id' -ne "")
     {

        $id = $i.'Issue Id'.replace("#","")
        $time = $i.'Time'.split(":");
        $time = [int]$time[0] + ([float]$time[1]/60)

        incrementWorklog $id $time $i.'User' $i.'Day' $i.'Time Entry'
     }
   }

}
