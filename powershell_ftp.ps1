Import-Csv .\Desktop\TestFTP\ftp.csv -Delimiter ";" | ForEach-Object {
    $Error.Clear()
    
    # csv
    $folder_name = $_.name
    $ftp_folder = $_.ftp_folder
    $ftp_port = $_.ftp_port
    $ftp_user = $_.ftp_user
    $ftp_password = $_.ftp_password
    $ftp = "ftp://" + $_.ftp + ":" + $ftp_port

    # Ftp folder empty
    if (!$ftp_folder)
    {
        $ftp_uri = $ftp + "/"
    }
    else
    {
        $ftp_uri = $ftp + "/" + $ftp_folder + "/"
    }

    $LocalPath = "FTPFiles\$folder_name\"

    # Check and create a folder
    if(!(Test-Path $LocalPath)){ New-Item -ItemType Directory -Path $LocalPath }

    # List of files from FTP
    function Get-FtpDir ($url, $credentials)
    {
    $request = [Net.FtpWebRequest]::Create($url)
    if ($credentials) { $request.Credentials = $credentials }
    $request.Method = [System.Net.WebRequestMethods+FTP]::ListDirectory
    (New-Object IO.StreamReader $request.GetResponse().GetResponseStream()).ReadToEnd() -split "`r`n"
    }
    
    # Needed for the download
    $webclient = New-Object System.Net.WebClient 
    $webclient.Credentials = New-Object System.Net.NetworkCredential($ftp_user,$ftp_password)  
    $webclient.BaseAddress = $ftp_uri

    # File filtering
    $files = Get-FTPDir $ftp_uri $webclient.Credentials | Where-Object { $_ -like "*.xls" -or $_ -like "*.xlsx" }
    # Download and delete file from ftp
    $files | ForEach-Object { 
    $webClient.DownloadFile( $_,$($LocalPath + $_) ) 
    if (Test-Path ($LocalPath+$_)) { Remove-FtpFile ($ftp_uri+$_) $webclient.Credentials }

    # Log
    Write-Output "$(Get-Date) $folder_name $ftp_uri)" + $Error[0] | Out-File "log.txt" -Append
}
