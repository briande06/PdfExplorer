
[System.Reflection.Assembly]::LoadFrom($PWD.Path + "\itextsharp.dll")


$workingdir = "K:\Papers"
$pdfs = Get-ChildItem $workingdir -File -Filter *.pdf

$table = "<table 
style=""width:100%"">
  <tr>
    <th>File</th>
    <th>Title</th>
    <th>Author</th>
    <th>Number of Page</th>
  </tr>"
  
foreach ($file in $pdfs) {
    $filename = [System.IO.Path]::GetFileNameWithoutExtension($file)
    $reader = New-Object iTextSharp.text.pdf.pdfreader $workingdir\$file
    if ($reader.Info.Title -and $reader.Info.Title -ne "") {
        $title = $reader.Info.Title.ToString()
    }
    else {
        $title = "Unknown"
    }

    if ($reader.Info.Author -and $reader.Info.Author -ne "") {
        $author = $reader.Info.Author.ToString()
    }
    else {
        $author = "Unknown"
    }
    #$relativeFile = Resolve-Path -Path $file.FullName -Relative
    $table = $table + "<tr> <td> <a href="".\" + $file.Name + " "" >" + $filename + "</a></td>  
          <td>" + $title + "</td> 
          <td>" + $author + "</td> 
          <td>" + $reader.NumberOfPages + "</td> 
          </tr>"
}
$table = $table + "</table>"

$style = "<style>

table, th, td {    border: 1px solid black;    border-collapse: collapse; }
th, td {    padding: 5px; }
th {    text-align: center; }
td:nth-child(1) {  font-weight:bold }


.footer {
position: absolute;
bottom: 0; margin: 0 auto;
}


body {
  text-align: center;
  margin: 10px auto;
  padding: 0 30px;
  max-width: 900px;
  font-family: 'Comfortaa', cursive;
 }

  h2 {
    color: red;
    text-shadow: 3px 3px 2px #000;
    text-align: center;
    font-size: 30pt;}
  
</style>"
$text = "<!DOCTYPE html> 
<html>"
$text += "<head> <link href=""https://fonts.googleapis.com/css?family=Comfortaa""rel=""stylesheet"">
" + $style + " 
<title> Papers </title>
</head>
 <body>"
$text += "<h2>Papers</h2>" 
$text += $table
$text += "</body>
 </html>"
$text > 'pdfs.html'