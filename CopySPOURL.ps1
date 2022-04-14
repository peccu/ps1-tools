$OutputEncoding = [console]::OutputEncoding;

# Get-Clipboard -Format Text -TextFormatType Html

function generatehtml($url, $anchor){
    $startFragment = 141
    $endfragment = $url.Length + $anchor.Length + 15 + $startFragment
    $endhtml = $endfragment + 36


    'Version:0.9
StartHTML:0000000105
EndHTML:{1:d10}
StartFragment:0000000141
EndFragment:{0:d10}
<html>
<body>
<!--StartFragment--><a href="{2:s}">{3:s}</a><!--EndFragment-->
</body>
</html>' -f $endfragment, $endhtml, $url, $anchor
}


function copyboth($html, $text){
    $data = New-Object System.Windows.Forms.DataObject
    $data.SetData([System.Windows.Forms.DataFormats]::Html, $html)
    $data.SetData([System.Windows.Forms.DataFormats]::Text, $text)
    [System.Windows.Forms.Clipboard]::SetDataObject($data)
}



function encodespourl($header, $target){
    Add-Type -AssemblyName System.Web
    $url = '{0:s}/{1:s}' -f $header, [System.Web.HttpUtility]::UrlEncode($target)
    $url = $url -replace '%2f', '/'
    $url = $url -replace '\+', '%20'
    return $url
}

$spolist = @{
    "C:\\Users\\kentaro.shimatani\\Accenture\\\[NTD\]クラウドコンシェ活動 - General" `
      = "/r/sites/NTD218/Shared%20Documents/General"
    "C:\\Users\\kentaro.shimatani\\Accenture\\【Nitto】Delivery Management - 130-ComputerResourceFinance" `
      = "/r/sites/NittoDeliveryManagement/Shared%20Documents/130-ComputerResourceFinance"
    "C:\\Users\\kentaro.shimatani\\Accenture\\NCPインフラ - NCPインフラ運用" `
      = "/r/sites/NTD_377/Shared%20Documents/NCP%E3%82%A4%E3%83%B3%E3%83%95%E3%83%A9%E9%81%8B%E7%94%A8"
    "C:\\Users\\kentaro.shimatani\\Accenture\\Nitto Account Leadership - コンピュータリソース" `
      = "/r/sites/NittoAccountLeadership/Shared%20Documents/General/%E3%81%9D%E3%81%AE%E4%BB%96/%E3%82%B3%E3%83%B3%E3%83%94%E3%83%A5%E3%83%BC%E3%82%BF%E3%83%AA%E3%82%BD%E3%83%BC%E3%82%B9"
    "C:\\Users\\kentaro.shimatani\\Accenture\\Nitto Account Leadership - Wave4-2安定化活動" `
      = "/r/sites/NittoAccountLeadership/Shared%20Documents/General/%E3%81%9D%E3%81%AE%E4%BB%96/Wave4-2%E5%AE%89%E5%AE%9A%E5%8C%96%E6%B4%BB%E5%8B%95"
    "C:\\Users\\kentaro.shimatani\\Accenture\\Nitto Account Leadership - Wave4-2" `
      = "/r/sites/NittoAccountLeadership/Shared%20Documents/General/%E6%8F%90%E6%A1%88%E6%B4%BB%E5%8B%95/Cloud%E9%96%A2%E9%80%A3/Wave4-2"
}


function typecode($targetpath){
  $isDirectory = (Get-Item -LiteralPath $targetpath) -is [System.IO.DirectoryInfo]
  if($isDirectory){ return ':f:' }

  $ext = (Get-Item -LiteralPath $targetpath).Extension
  if($ext -eq '.pptx'){ return ':p:'}
  if($ext -eq '.txt'){ return ':t:'}
  if($ext -eq '.xlsx'){ return ':x:'}
  if($ext -eq '.pdf'){ return ':b:'}
  # msg, zip
  # maybe universal
  return ':u:'
}

function pathtourl($targetpath){
    $type = typecode $targetpath
    # write-Output 'type', $type
    $spolist.keys | ForEach-Object {
        $message = 'matching {0}' -f $_
        # Write-Output $message
        if ($targetpath -match ('^{0}' -f $_)){
            $header = 'https://ts.accenture.com/{0}{1}' -f $type, $spolist[$_]
            $message = 'found {1} for {0}' -f $header, $_
            # Write-Output $message
            $target = $targetpath -replace $_, ''
            $target = $target -replace '^\\', ''
            $target = $target -replace '\\', '/'
            # Write-Output $target
            $url = encodespourl $header $target
            $url = '{0}?web=1' -f $url
            # Write-Output $url
            return $url
        }
    }
}
# "C:\Users\kentaro.shimatani\Accenture\[NTD]クラウドコンシェ活動 - General\900_その他\920_ACP\20200831_ACP_X_References\acp-x-updates.xlsx"
# "C:\Users\kentaro.shimatani\Accenture\【Nitto】Delivery Management - 130-ComputerResourceFinance\References\20220414_ACP4-ACPX\Computer Res.ACP4vsACPX_E_SC.pptx"

# $t = "C:\Users\kentaro.shimatani\Accenture\【Nitto】Delivery Management - 130-ComputerResourceFinance\MME-Forecast\Nitto ITO_Financial Status_Comp Res_20220329_v1.0.xlsx"
# $t = "C:\Users\kentaro.shimatani\Accenture\【Nitto】Delivery Management - 130-ComputerResourceFinance\管理プロセス"
# # https://ts.accenture.com/:f:/r/sites/NittoDeliveryManagement/Shared%20Documents/130-ComputerResourceFinance/%E7%AE%A1%E7%90%86%E3%83%97%E3%83%AD%E3%82%BB%E3%82%B9
# # https://ts.accenture.com/:f:/r/sites/NittoDeliveryManagement/Shared%20Documents/130-ComputerResourceFinance/%e7%ae%a1%e7%90%86%e3%83%97%e3%83%ad%e3%82%bb%e3%82%b9
# $t = "C:\Users\kentaro.shimatani\Accenture\【Nitto】Delivery Management - 130-ComputerResourceFinance\管理プロセス\コンピュータリソース対応プロセス_v20220317.pptx"
# # https://ts.accenture.com/:p:/r/sites/NittoDeliveryManagement/Shared%20Documents/130-ComputerResourceFinance/%E7%AE%A1%E7%90%86%E3%83%97%E3%83%AD%E3%82%BB%E3%82%B9/%E3%82%B3%E3%83%B3%E3%83%94%E3%83%A5%E3%83%BC%E3%82%BF%E3%83%AA%E3%82%BD%E3%83%BC%E3%82%B9%E5%AF%BE%E5%BF%9C%E3%83%97%E3%83%AD%E3%82%BB%E3%82%B9_v20220317.pptx
# # https://ts.accenture.com/:f:/r/sites/NittoDeliveryManagement/Shared%20Documents/130-ComputerResourceFinance/%e7%ae%a1%e7%90%86%e3%83%97%e3%83%ad%e3%82%bb%e3%82%b9/%e3%82%b3%e3%83%b3%e3%83%94%e3%83%a5%e3%83%bc%e3%82%bf%e3%83%aa%e3%82%bd%e3%83%bc%e3%82%b9%e5%af%be%e5%bf%9c%e3%83%97%e3%83%ad%e3%82%bb%e3%82%b9_v20220317.pptx
# pathtourl $t
# https://ts.accenture.com/:p:/r/sites/NittoDeliveryManagement/Shared%20Documents/130-ComputerResourceFinance/References/20220414_ACP4-ACPX/Computer%20Res.ACP4vsACPX_E_SC.pptx?d=w68324e729c2c4945acb431ed1495e9d7&csf=1&web=1&e=JDKnSE
# https://ts.accenture.com/:x:/r/sites/NittoDeliveryManagement/Shared%20Documents/130-ComputerResourceFinance/MME-Forecast/Computer%20Res.Forecast_ACPX_20220405_v1.xlsx?d=w06cd3fa37f914f4d94dda5cb7ba581ee&csf=1&web=1&e=0IwA76
# https://ts.accenture.com/:t:/r/sites/NittoDeliveryManagement/Shared%20Documents/130-ComputerResourceFinance/MME-Forecast/memo.txt?csf=1&web=1&e=wTqixu
# https://ts.accenture.com/:u:/r/sites/NTD218/Shared%20Documents/General/900_%E3%81%9D%E3%81%AE%E4%BB%96/920_ACP/20200831_ACP_X_References/RE%20Resell%20model%20for%20Nitto%20Denko.msg?csf=1&web=1&e=EcHUOF
# https://ts.accenture.com/:b:/r/sites/NTD218/Shared%20Documents/General/900_%E3%81%9D%E3%81%AE%E4%BB%96/920_ACP/20200831_ACP_X_References/What%27s%20New%20in%20ACP%20X%20112020.pdf?csf=1&web=1&e=iRrJXs
# https://ts.accenture.com/:u:/r/sites/NTD218/Shared%20Documents/General/900_%E3%81%9D%E3%81%AE%E4%BB%96/920_ACP/20200831_ACP_X_References/20200831_ACP_X_References.zip?csf=1&web=1&e=iY9VKr



function extractanchor($targetpath){
    $isDirectory = (Get-Item -LiteralPath $targetpath) -is [System.IO.DirectoryInfo]
    if($isDirectory){
        return '{0}' -f (Get-Item -LiteralPath $targetpath).BaseName
    }else{
        return '{0}{1}' -f (Get-Item -LiteralPath $targetpath).BaseName, (Get-Item -LiteralPath $targetpath).Extension
    }
}
# extractanchor $t | Write-Output
# $t = "C:\Users\kentaro.shimatani\Accenture\【Nitto】Delivery Management - 130-ComputerResourceFinance\管理プロセス"
# extractanchor $t | Write-Output


# $targetpath=(Get-Clipboard -Format Text) -replace '^"|"$', ''
# write-output $targetpath
# $url = pathtourl $targetpath
# $anchor = extractanchor $targetpath
# '{0} -> {1}' -f $anchor, $url | write-output

function main(){
    $targetpath=((Get-Clipboard -Format Text) -replace '^"', '') -replace '"$', ''
    if("" -eq $targetpath -or -Not (Test-Path -LiteralPath "$targetpath")){
        Write-Output "not path", $targetpath
        return
    }
    # write-output "it is path"
    $url = pathtourl $targetpath
    $anchor = extractanchor $targetpath

    # $url = 'https://ts.accenture.com/:f:/r/sites/NittoDeliveryManagement/Shared%20Documents/130-ComputerResourceFinance/MME-Forecast?csf=1&amp;web=1&amp;e=JwyrIj'
    # $anchor = 'MME-Forecast'
    # generatehtml $url $anchor | Set-Clipboard -AsHtml


    # $targetpath = "C:\Users\kentaro.shimatani\Accenture\【Nitto】Delivery Management - 130-ComputerResourceFinance\MME-Forecast\Nitto ITO_Financial Status_Comp Res_20220329_v1.0.xlsx"
    # $target = transferpath $targetpath

    # $header = 'https://ts.accenture.com/:f:/r/sites/'
    # $target = 'NittoAccountLeadership/Shared Documents/General/その他/コンピュータリソース/課金管理/月次利用明細/提出用/'
    # $url = encodespourl $header $target
    # $url

    # NittoAccountLeadership/Shared Documents/General/その他/コンピュータリソース/課金管理/月次利用明細/提出用/

    # https://ts.accenture.com/:f:/r/sites/NittoAccountLeadership/Shared%20Documents/General/%E3%81%9D%E3%81%AE%E4%BB%96/%E3%82%B3%E3%83%B3%E3%83%94%E3%83%A5%E3%83%BC%E3%82%BF%E3%83%AA%E3%82%BD%E3%83%BC%E3%82%B9/%E8%AA%B2%E9%87%91%E7%AE%A1%E7%90%86/%E6%9C%88%E6%AC%A1%E5%88%A9%E7%94%A8%E6%98%8E%E7%B4%B0/%E6%8F%90%E5%87%BA%E7%94%A8?csf=1&web=1&e=2BkQtd
    # https://ts.accenture.com/:f:/r/sites/NittoAccountLeadership/Shared%20Documents/General/%E3%81%9D%E3%81%AE%E4%BB%96/%E3%82%B3%E3%83%B3%E3%83%94%E3%83%A5%E3%83%BC%E3%82%BF%E3%83%AA%E3%82%BD%E3%83%BC%E3%82%B9/%E8%AA%B2%E9%87%91%E7%AE%A1%E7%90%86/%E6%9C%88%E6%AC%A1%E5%88%A9%E7%94%A8%E6%98%8E%E7%B4%B0/%E6%8F%90%E5%87%BA%E7%94%A8
    # $url = 'https://ts.accenture.com/sites/NittoAccountLeadership/Shared Documents/General/その他/コンピュータリソース/課金管理/月次利用明細/提出用/'


    $html = generatehtml $url $anchor
    copyboth $html $url
    write-output "path copied"
}
main
