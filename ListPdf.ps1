
[System.Reflection.Assembly]::LoadFrom($PWD.Path+"\itextsharp.dll")


$workingdir=$PWD.Path+"\..\Papers"
$pdfs = Get-ChildItem $workingdir -File -Filter *.pdf


$style = "<style> 
table, th, td {    border: 1px solid black;    border-collapse: collapse; }
th, td {    padding: 5px; }
th {    text-align: left; }
td:nth-child(1) {  font-weight:bold }
</style>"

$table = "<table 
style=""width:100%"">
  <tr>
    <th>File</th>
    <th>Title</th>
    <th>Author</th>
  </tr>"
  
foreach($file in $pdfs)
{
  $filename = [System.IO.Path]::GetFileNameWithoutExtension($file)
  $reader = New-Object iTextSharp.text.pdf.pdfreader $workingdir\$file
  if($reader.Info.Title -and $reader.Info.Title -ne "")
  {
    $title = $reader.Info.Title.ToString()
  }
  else {
    $title = "Unknown"
  }

  if($reader.Info.Author -and $reader.Info.Author -ne "")
  {
    $author = $reader.Info.Author.ToString()
  }
  else {
    $author = "Unknown"
  }
  $relativeFile=Resolve-Path -Path $file.FullName -Relative
  $table = $table + "<tr> <td> <a href=""" +$relativeFile + " "" >" + $filename + "</a></td>  
          <td>" + $title + "</td> 
          <td>"+ $author +"</td> 
          </tr>"
}
$table += "</table>"


$text = "<!DOCTYPE html> 
<html>"
$text += "<head>" + $style + "</head> <body>"
$text += "<h2>Summary</h2>"
$text += $table
$text += "</body> </html>"

$text > 'pdfs.html'