#Создаем переменную с текущей датой
$dt=Get-Date -Format "dd-MM-yyyy_HH-mm"

#Создаем каталог для логов, если его нет
New-Item -ItemType directory log -Force | out-null

#Путь с датой в название файла
$logfile="C:\Scripts\log\pdf\"+$dt+"_LOG.log"

#Начинаем записывать логи
Start-Transcript -Path $Logfile -Append

#Read json file
$j = get-content "C:\Scripts\base.json" -Encoding UTF8 | ConvertFrom-Json

#Find files in a folder 
$files_pdf=get-childitem -Path "D:\Shares\Договоры какие-то сканы" –recurse | where-object {$_.Creationtime -gt (get-date).AddMinutes(-60) -and ($_.extension -eq ".pdf") -and ($_.Name -like "*оговор*")} | Foreach-Object { $_.FullName }

#Erasing data in the array
$arr_fl4 = @()

#Iterating through the data in the array
$arr_fl4 = foreach ($item in $files_pdf)
{
    if ($item -match "(?'path'(?<=\\\w.....\\).+(?=\\.*\.[PDF|pdf]))")
    {
        $path = $($Matches['path'])

        if ($item -match "(?'name'([^\\]+?)(?=\.[PDF|pdf]))")
        {
           $name = $($Matches['name'])
        }

     }

        [PSCustomObject]@{
            Name = $name
            Path = $path
        }
}
    
$Header = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">
<head>
<title></title>
<style type="text/css">
table a:link {
	color: #666;
	font-weight: bold;
	text-decoration:none;
}
table a:visited {
	color: #999999;
	font-weight:bold;
	text-decoration:none;
}
table a:active,
table a:hover {
	color: #bd5a35;
	text-decoration:underline;
}
table {
	font-family:Arial, Helvetica, sans-serif;
	color:#666;
	font-size:12px;
	text-shadow: 1px 1px 0px #fff;
	background:#eaebec;
	margin:20px;
	border:#ccc 1px solid;

	-moz-border-radius:3px;
	-webkit-border-radius:3px;
	border-radius:3px;

	-moz-box-shadow: 0 1px 2px #d1d1d1;
	-webkit-box-shadow: 0 1px 2px #d1d1d1;
	box-shadow: 0 1px 2px #d1d1d1;
}
table th {
	padding:21px 25px 22px 25px;
	border-top:1px solid #fafafa;
	border-bottom:1px solid #e0e0e0;

	background: #ededed;
	background: -webkit-gradient(linear, left top, left bottom, from(#ededed), to(#ebebeb));
	background: -moz-linear-gradient(top,  #ededed,  #ebebeb);
}
table th:first-child {
	text-align: center;
	padding-left:20px;
}
table tr:first-child th:first-child {
	-moz-border-radius-topleft:3px;
	-webkit-border-top-left-radius:3px;
	border-top-left-radius:3px;
}
table tr:first-child th:last-child {
	-moz-border-radius-topright:3px;
	-webkit-border-top-right-radius:3px;
	border-top-right-radius:3px;
}
table tr {
	text-align: center;
	padding-left:20px;
}
table td:first-child {
	text-align: left;
	padding-left:20px;
	border-left: 0;
}
table td:last-child {
	text-align: left;
	padding-left:20px;
	border-left: 0
}
table td {
	padding:18px;
	border-top: 1px solid #ffffff;
	border-bottom:1px solid #e0e0e0;
	border-left: 1px solid #e0e0e0;

	background: #fafafa;
	background: -webkit-gradient(linear, left top, left bottom, from(#fbfbfb), to(#fafafa));
	background: -moz-linear-gradient(top,  #fbfbfb,  #fafafa);
}
table tr.even td {
	background: #f6f6f6;
	background: -webkit-gradient(linear, left top, left bottom, from(#f8f8f8), to(#f6f6f6));
	background: -moz-linear-gradient(top,  #f8f8f8,  #f6f6f6);
}
table tr:last-child td {
	border-bottom:0;
}
table tr:last-child td:first-child {
	-moz-border-radius-bottomleft:3px;
	-webkit-border-bottom-left-radius:3px;
	border-bottom-left-radius:3px;
}
table tr:last-child td:last-child {
	-moz-border-radius-bottomright:3px;
	-webkit-border-bottom-right-radius:3px;
	border-bottom-right-radius:3px;
}
table tr:hover td {
	background: #f2f2f2;
	background: -webkit-gradient(linear, left top, left bottom, from(#f2f2f2), to(#f0f0f0));
	background: -moz-linear-gradient(top,  #f2f2f2,  #f0f0f0);	
}
</style>
"@    
Write-Host $arr_fl4

function Send-Email {
    param(
        $emailList,
        $arr_docs
    )
    # Sender and Recipient Info
    $MailFrom = "service@kontora1.com"

    # Sender Credentials
    $Username = "service"
    $Password = "password"

    # Server Info
    $SmtpServer = "mail.kontora1.com"

    # Message stuff
    $MessageSubject = "Новые договора" 

    # Construct the SMTP client object, credentials, and send
    $Smtp = New-Object Net.Mail.SmtpClient($SmtpServer)
    $Smtp.Credentials = New-Object System.Net.NetworkCredential($Username,$Password)

    write-host "---------------------------------"
    write-host $emailList
    write-host $arr_docs

    foreach ($MailTo in $emailList) {
        $a = $arr_docs.count
        $output = $arr_docs | ConvertTo-Html -Property Name,Path -Head $Header -pre "<span>Размещены новые документы, $a шт. <span>" | foreach {
        $PSItem -replace "<th>Name</th>", "<th>Название документа</th>" -replace "<th>Path</th>","<th>Путь</th>"}
        $Message = New-Object System.Net.Mail.MailMessage $MailFrom,$MailTo
        $Message.IsBodyHTML = $true
        $Message.Subject = $MessageSubject
        $Message.Body = $output
        $Smtp.Send($Message)
    }    
}

if ($files_pdf.count -gt 0) {
    foreach($item in $j) {
        $arr_files = @()
        foreach($element in $arr_fl4){
            if($element.Path -like $item.bits -And $element.Name -ne $null)
            {   
                $arr_files += $element                
             }
         }
        if($arr_files.count -gt 0)  
        {
            Send-Email -emailList $item.emails -arr_docs $arr_files
        }
        
    }
}

#Удаляем файлы старше 30 дней
$Folder = "C:\Scripts\log\pdf\"
Get-ChildItem $Folder -Recurse -Force -ea 0 |
? {!$_.PsIsContainer -and $_.LastWriteTime -lt (Get-Date).AddDays(-30)} |
ForEach-Object {
   $_ | del -Force
   $_.FullName
}

#Заканчиваем записывать логи
Stop-Transcript