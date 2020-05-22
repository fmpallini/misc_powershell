#Example of use qTorrent, run after complete download: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -executionpolicy unrestricted -file "C:\Users\fmpal\Documents\copy_file.ps1" "%F"

#config
$max_retry = 3
$destination_path = "\\192.168.1.1\filmes (at My_Passport)"
$movie_extensions = @(".avi",".mkv",".mp4")
$min_size = 100000000 # 100mb

#begin
$downloadPath = $Args[0]
$counter = 0;

#check if is movie
$literalPath = "\\?\"+$downloadPath
$dir = Get-ChildItem -LiteralPath $literalPath -Recurse
$isMovie = $false

For ($i=0; $i -lt $movie_extensions.length; $i++) {
	 If($dir | Where {$_.Name.ToLower().EndsWith($movie_extensions[$i]) -and $_.Length -gt $min_size})
	 {
		$isMovie = $true
		Break;
	 }
}

If(!$isMovie)
{
	Exit
}

#try to reach the router
While($counter -Lt $max_retry -And -Not(Test-Path -LiteralPath $destination_path))
{
    $counter++
    Start-Sleep -Seconds 1
}

If(Test-Path -LiteralPath $destination_path)
{
    Copy-Item -LiteralPath $literalPath -Recurse $destination_path
    explorer $destination_path
}
Else
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [System.Windows.Forms.Messagebox]::Show("Something wrong happened. Destination: " + $destination_path + " Original File: " + $downloadPath)
}