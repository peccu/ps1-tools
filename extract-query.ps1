param ([Parameter(Mandatory)]$filepath)

#$filepath = "C:\Users\kentaro.shimatani\Codes\LB-flowlogs\PowerQuery.xlsx"
#$filepath = ".\Azureñæç◊çÏê¨çÏã∆_withCloudHealth_202110.xlsx"
$filepath = "C:/Users/kentaro.shimatani/Downloads/query-testing foo bar/Nitto ITO_Financial Status_Comp Res_20220404_v1.0.xlsx"
Write-Output $filepath
Get-Item $filepath
#Throw "foo"
# Write-Output $args[1]
# $filepath = $args[1]
# if ($filepath -eq ""){
# Throw "no file specified"
# }
# $closeExcel = $true

function createExcel(){
    $xl = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    if ( $xl.ActiveWorkbook.Name.length -eq 0 ){
        return New-Object -ComObject Excel.Application
    }
    return $xl
}

function openExcel{
    param ( $excel, $file )
    $excel.Visible = $true
    # todo open or pickup
    $book =  $excel.Workbooks | ? {$_.FullName -eq "$file"}
    if ($book) {
        Write-Output "found file. use existed"
        # $closeExcel = $false
        return $book
    }
    Write-Output "not found in current process. open it"
    return $excel.Workbooks.Open($file)
    # $excel.save
    # return $targetbook
}

# https://stackoverflow.com/a/28237896
Function pause ($message)
{
    # Check if running Powershell ISE
    if ($psISE)
    {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("$message")
    }
    else
    {
        Write-Host -NoNewline "$message" -ForegroundColor Yellow
        $x = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

function extractQueries($book, $outdir){
    if (test-path $outdir){
        Write-Output "Already exists the same name directory. Using it."
    }else{
        New-Item -ItemType Directory -Path $outdir
    }

    # create index file
    $indexfile = (Join-Path $outdir "index.txt")
    if (-Not (Test-Path $indexfile)){
        New-Item $indexfile -ItemType File
        Start-Sleep -s 1
    }
    Set-Content -Path $indexfile `
      -Value ("# Queries size: " + $book.Queries.Length)

    # generate query files
    Write-Output "Queries size: " $book.Queries.Length
    $book.Queries | foreach {$i=0} {
        write-output $_.Name
        Write-Output $i
        $filename = "Query_" + [string]$i + ".m"
        $outfile = (Join-Path $outdir $filename)

        # append to index
        Add-Content -Path $indexfile `
          -Value ("- " + $_.Name + ": " + $filename)
        if($_.Description -ne ""){
            Add-Content -Path $indexfile -Value $_.Description
        }

        #$of = (get-item -literalpath (gc $outfile))
        #Write-Output $of
        #continue
        if (-Not (Test-Path $outfile)){
            New-Item $outfile -ItemType File
            Start-Sleep -Milliseconds 500
        }
        Set-Content -Path $outfile `
          -Value ($_.Name + [Environment]::NewLine + $_.Formula)
        $i++
    }
    Write-Output "finished extract."
}

function importQueries($book, $outdir){
    if (-Not (test-path $outdir)){
        Write-Output "No directory."
        return
    }

    $queries = Get-ChildItem -Path (join-path $outdir "*.m") -Name
    for ($i=0; $i -lt $queries.count; $i++){
        #$queries | % {$i=0} {
        $filename = "Query_" + [string]$i + ".m"
        $content = Get-Content -Path (join-path $outdir $filename)
        $qname = $content | Select-Object -First 1
        $qformula = [string]::Join("`r`n`t", ($content | Select-Object -Skip 1))
        Write-Output $qname
        Write-Output $i
        $book.Queries[$i].Name = $qname
        $book.Queries[$i].Formula = $qformula
    }
}

function main($filepath){
    echo "opening $filepath"
    if (! [System.IO.Path]::IsPathRooted("../Scripts")){}
    Write-Output (get-item -LiteralPath $filepath)
    $dirname = (get-item $filepath ).DirectoryName
    $filename = (get-item $filepath ).Name
    $basename = (get-item $filepath ).BaseName
    $outdir = Join-Path $dirname $basename
    $abspath = Join-Path $dirname $filename
    try{
        $excel = createExcel
        $book = openExcel $excel $abspath
        sleep 5

        echo "extracting"
        extractQueries $book $outdir
        Write-Output "extraction completed. Please edit queries freely. files are in"
        Write-Output $outdir
        pause("Please hit any key to save modified queries...")
        importQueries $book $outdir
        Write-Output "imported."
        # Write-Output "imported. now saving"
        # $book.Save()
        Write-Output "finished."
        # if($closeExcel){
        #     Write-Output "Closing Excel"
        #     $excel.Quit()
        # }
    }catch{
        Write-Output "something error:"
        Write-Output $_
        # if($closeExcel){
        #     Write-Output "Closing Excel"
        #     $excel.Quit()
        # }
    }

}


main $filepath
