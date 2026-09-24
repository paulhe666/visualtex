param([int]$WordProcessId,[string]$Label='scan-mtplaceref')
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$word=Connect-RunningWord $WordProcessId
try {
  $doc=$word.ActiveDocument
  $rows=@()
  for($i=1;$i -le $doc.Fields.Count;$i++) {
    $field=$null;$code=$null;$result=$null;$range=$null
    try {
      $field=$doc.Fields.Item($i)
      if($field.Type -ne 51){ continue } # wdFieldMacroButton
      $code=$field.Code
      $codeText=$code.Text
      if($null -eq $codeText){$codeText=''}
      if($codeText.IndexOf('MACROBUTTON MTPlaceRef',[StringComparison]::OrdinalIgnoreCase) -lt 0){ continue }
      $result=$field.Result
      $fieldStart=$code.Start-1
      $codeBoundary=$doc.Range($code.End,$code.End+1).Text
      if($codeBoundary -eq [string][char]0x15){$fieldEnd=$code.End+1}
      elseif($codeBoundary -eq [string][char]0x14 -and $doc.Range($result.End,$result.End+1).Text -eq [string][char]0x15){$fieldEnd=$result.End+1}
      else{throw "invalid field boundary at $($code.Start)"}
      $range=$doc.Range($fieldStart,$fieldEnd)
      $xmlText=$range.WordOpenXML
      $xml=New-Object System.Xml.XmlDocument
      $xml.PreserveWhitespace=$true
      $xml.LoadXml($xmlText)
      $ns=New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
      $ns.AddNamespace('pkg','http://schemas.microsoft.com/office/2006/xmlPackage')
      $ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
      $paras=$xml.SelectNodes("//pkg:part[@pkg:name='/word/document.xml']/pkg:xmlData/w:document/w:body/w:p",$ns)
      $ownerCount=0
      foreach($p in $paras){
        $instruction=''
        foreach($n in $p.SelectNodes('.//w:instrText',$ns)){$instruction += $n.InnerText}
        if($instruction.IndexOf('MACROBUTTON MTPlaceRef',[StringComparison]::OrdinalIgnoreCase) -ge 0){$ownerCount++}
      }
      $row=[pscustomobject]@{FieldIndex=$i;CodeStart=$code.Start;FieldStart=$fieldStart;FieldEnd=$fieldEnd;XmlLength=$xmlText.Length;Paragraphs=$paras.Count;OwnerParagraphs=$ownerCount}
      $rows += $row
      if($ownerCount -ne 1){
        $path=Join-Path $PSScriptRoot "evidence\$Label-field-$i.xml"
        [IO.File]::WriteAllText($path,$xmlText,[Text.UTF8Encoding]::new($false))
      }
    } finally {
      foreach($obj in @($range,$result,$code,$field)){if($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj)}}
    }
  }
  [pscustomobject]@{Document=$doc.Name;Total=$rows.Count;Bad=@($rows|Where-Object {$_.OwnerParagraphs -ne 1}).Count;Rows=$rows}|ConvertTo-Json -Depth 5
} finally {
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
