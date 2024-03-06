# This script is to be primarily used with Audiobooks, will do the following.
# 1. Create seperate NFO files using the metadata from the first audio file scanned in each directory.
# 2. If an CUE file exists in the same folder as a M4B it will update the file path in the CUE file if needed and count the numbr or "Chapter" references. 
# 3. Will create files Reader.txt (Author) and Desc.txt (comments) using metadata from audiofiles
# 4. If a folder.jpg exist it will copy to cover.jpg and if folder.jpg width is larger 800 it will copy to the name of the "<artist> - <album>.HR.jpg"    

##########
# PreReqs
##########
# Requires TagLibSharp.dll in the same directory as the script, prebuilt DLL grab from: https://github.com/mono/taglib-sharp
# Requires Tone.exe in the same directory as the script, grab from: https://github.com/sandreas/tone
# Requires Powershell module MediaInfo, install module from Admin PoSH run: Install-Module -name get-mediainfo

########
# Usage
########
# Drag parent folder that has multiple subfolders containing audio files onto the BAT file (named the same as the PS1 script)

[CmdletBinding()]
param (
    [Parameter(ValueFromRemainingArguments=$true)]
    $Path
)

$scriptpath = $MyInvocation.MyCommand.path
$dir = Split-Path $scriptpath
Push-Location $dir

#$scriptDirectory = (Get-Item $PSCommandPath).DirectoryName
$taglibsharp = $dir + '\TagLibSharp.dll'

[Reflection.Assembly]::LoadFile($taglibsharp)
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
# [System.Windows.Forms.MessageBox]::Show("-p0:$($pubdate[0]) $([System.Environment]::NewLine) -p1:$($pubdate[1]) $([System.Environment]::NewLine) -p2:$($pubdate[2]) $([System.Environment]::NewLine) -p3:$($pubdate[3]) $([System.Environment]::NewLine)")

##### Enable to delete existing Reader.txt desc.txt files
#Get-childitem -Path "$path\*" -Recurse -include reader.txt | Remove-Item -Force
#Remove-Item "$folder\reader.txt" -Force

#[System.Windows.Forms.MessageBox]::Show("Path =  $Path ")

Function Get-AudioDetails {
    param ($targetDirectory)

    Get-ChildItem -LiteralPath $folder -recurse | Where-Object {$_.name -match $pattern}| ForEach-Object {
        $Audio = [TagLib.File]::Create($_.FullName)

        New-Object PSObject -Property @{
            Name = $_.FullName
            Duration = $Audio.Properties.Duration.TotalMinutes
		}
    }
}

$pattern = "\.(mp4|mp3|m4b|m4a)$"

#$targetDir = (Get-ChildItem  -LiteralPath $path -Recurse -Directory).FullName

# Manual updates uncomment
#$path = "L:\@Audiobooks\<Author>"
#$targetDir = $path

$targetDir = Get-ChildItem -LiteralPath $path -Recurse | Where-Object{ $_.PSIsContainer } | Select-Object FullName | ForEach-Object{$_.fullname}
 foreach ($folder in $targetDir){

   Get-ChildItem -LiteralPath $folder -recurse | Where-Object {$_.name -match $pattern} |
   Select-Object -First 1 |
   ForEach-Object {
   Write-Host "" `n
    #$imageHD = ""
    #$imageHDExists = ""
    $file = $_.FullName
    $fileShort = $_.Name
    $extension = [System.IO.Path]::GetExtension($fileShort)
    $folder = Split-Path -Path $file
    $FolderJPGExists = Test-Path  -LiteralPath "$folder\folder.jpg"
    $FileExists = Test-Path  -LiteralPath "$folder\reader.txt"
    $CoverJPGExists = Test-Path  -LiteralPath "$folder\cover.jpg"
    $fileNFOExist = (Get-ChildItem -Path $folder -force | Where-Object Extension -in ('.nfo') | Measure-Object).Count -ne 0

    #$backupNFO = "$folder\$fileNFO" + '.txt'
    #$backupNFOExists = Test-Path $backupNFO
    #$imageHRExists = Test-Path  -LiteralPath "$folder\$imageHR"

    $duration2    = ""
    $narrator     = ""
    $comments     = ""
    $artist       = ""
    $album        = ""
    $Channel      = ""
    $BitRate      = ""
    $BitRateMode  = ""
    $OvAllBitRate = ""
    $Encoded_Lib  = ""
    $SamplingRate = ""
    $Format       = ""
    $Format_Ver   = ""
    $Format_Prof  = ""
    $series2      = ""
    $seriesnum    = ""
    $asin         = ""
    $isbn         = ""
    $fileNFO      = ""
    $imageHR      = ""
    $chapterCount = ""
    
    ################
    # TagLib values
    ################
        $media = [TagLib.MPEG.File]::Create($file)

        $comments     = $media.Tag.comment
        $narrator     = $media.Tag.composers
        $artist       = $media.Tag.artists
        $albumartists = $media.Tag.albumartists
        $album        = $media.Tag.album
        $year         = $media.Tag.year
        #$genres       = $media.Tag.genres
        #$publisher    = $media.Tag.publisher
        $subtitle     = $media.Tag.subtitle
		$Duration	  = $media.Tag.Duration
        $JoinedGenres = $media.Tag.JoinedGenres
        $Copyright    = $media.Tag.Copyright
        $AmazonId     = $media.Tag.AmazonId #ASIN
        $isbn         = $media.Tag.isbn
        $bitrate = $media.Properties.AudioBitrate

    ##############
    # Tone values
    ##############
        $bJSON = & "$dir\tone.exe" dump "$file" --format json  --exclude-property=embeddedPictures | ConvertFrom-Json #--query "$.meta.additionalFields"

        # format agnostic
        $metaFormatTone     = $bJSON.audio.metaFormat    # mp4, id3V23, id3V1
        $metaFormatChoice   = $metaFormatTone[0]
        $encodingToolTone   = $bJSON.meta.encodingTool
        $genre              = $bJSON.meta.genre
        $subtitleTone       = $bJSON.meta.subtitle
        $commentTone        = $bJSON.meta.comment
        $descriptionTone    = $bJSON.meta.description
        $publisherTone      = $bJSON.meta.publisher
        [string]$publishYear= $bjson.meta.recordingDate.Year
        
        # MP3 specific
        $pubDateMP3Tone     = $bJSON.meta.additionalFields.tdrl
        $asinMP3Tone        = $bJSON.meta.additionalFields.asin
        $isbnMP3Tone        = $bJSON.meta.additionalFields.isbn
        $SeriesMP3Tone      = $bJSON.meta.additionalFields.SERIES
        $SeriesNumMP3Tone   = $bJSON.meta.additionalFields.'SERIES-PART'
        
        # MB4 specific
        $asinM4BTone        = $bJSON.meta.additionalfields.'----:com.apple.iTunes:ASIN'
        $isbnM4BTone        = $bJSON.meta.additionalfields.'----:com.apple.iTunes:ISBN'
        $SeriesM4bTone      = $bJSON.meta.additionalfields.'----:com.apple.iTunes:SERIES'
        $SeriesNumM4BTone   = $bJSON.meta.additionalfields.'----:com.apple.iTunes:SERIES-PART'
        $pubDateM4BTone     = $bJSON.meta.additionalfields.'----:com.apple.iTunes:RELEASETIME'

        $FormatTone         = $bJSON.audio.formatShort
        $Encoded_LibTone    = $bJSON.audio.format
        $overAllBitRateTone = $bJSON.audio.bitrate
        $samplingRateTone   = ($bJSON.audio.sampleRate)/1000
        $channelTone        = $bJSON.audio.channels.count
        $channelDescTone    = $bJSON.audio.channels.description
        $VBRTone            = $bJSON.audio.vbr

    ##################
    # Mediainf values
    ##################
        #$mediainf = Get-MediaInfo $file -Audio

        $Channel      = Get-MediaInfoValue $file -Kind Audio -Parameter 'Channel(s)'
        $BitRate      = Get-MediaInfoValue $file -Kind Audio -Parameter 'BitRate'
        $BitRateMode  = Get-MediaInfoValue $file -Kind General -Parameter 'OverallBitRate_Mode'
        $OvAllBitRate = Get-MediaInfoValue $file -Kind General -Parameter 'OverallBitRate'
        $Encoded_Lib  = Get-MediaInfoValue $file -Kind Audio -Parameter 'Encoded_Library'
        $SamplingRate = Get-MediaInfoValue $file -Kind Audio -Parameter 'SamplingRate'
        $Format       = Get-MediaInfoValue $file -Kind Audio -Parameter 'Format'              
        $Format_Ver   = Get-MediaInfoValue $file -Kind Audio -Parameter 'Format_Version'                 
        $Format_Prof  = Get-MediaInfoValue $file -Kind Audio -Parameter 'Format_Profile'
        $RecordedDate = Get-MediaInfoValue $file -Kind General -Parameter 'Recorded_Date'
            
        If($Encoded_Lib -match 'LAMEUUU'){$Encoded_Lib= $Encoded_Lib.Replace("LAMEUUU","")}
        #write-host "$Duration " -ForegroundColor Yellow
                
        $SamplingRate = $SamplingRate/1000
        $OvAllBitRate = $OvAllBitRate/1000
        
    ##############################################################################
    # NFO values based on MetaForma type, use combined TagLib + Tone + MediaInfo
    ##############################################################################
        if(($metaFormatChoice -eq "id3V23")-or($metaFormatChoice -eq "id3V1")){
            $series     = $SeriesMP3Tone
            if($null -ne $SeriesNumMP3Tone){$SeriesNum  = '#' + $SeriesNumMP3Tone}   
            $encoded_Lib = $encoded_LibTone
            $isbn = $isbnMP3Tone
            $asin = $asinMP3Tone
            $subtitle = $subtitleTone
            $encodingTool = $encodingToolTone
        }

        if($metaFormatChoice -eq 'mp4'){
            $series       = $SeriesM4bTone   
            $asin         = $asinM4BTone     
            #$publishDate  = $pubDateM4BTone
            #$comments     = $descriptionTone
            #$comments     = $commentTone.ReplaceLineEndings() #Tone comments use as TagLib has charcter limiation for mp4, but will mess up eoL
            if($null -ne $SeriesNumM4BTone){$SeriesNum  = '#' + $SeriesNumM4BTone}
            $encoded_Lib = $encoded_LibTone
            $isbn = $isbnM4BTone
            $subtitle = $subtitleTone
            $encodingTool = $encodingToolTone
        }

        if ($null -ne $album){
        $album = $album.Replace(';',',').Replace(':','')
        }
        
        if ($null -ne $JoinedGenres){
        $jGenre2 = $JoinedGenres.Replace('; ',', ').Replace('/',', ').Replace('Audiobook,','')
        $jGenre2 = $jGenre2.Trim()
        }

        if ($null -ne $Copyright){
            $Copyright = $Copyright.Replace('©','(C)').Replace('(c)','(C)').Replace('℗','(P)').Replace('(p)','(P)')

            $pubdate = $Copyright -Split("\(P\)")
            $pubdate2 = $pubdate[1]
            $pubdate3 = $pubdate2.Split(' ')
            $copyDate = $pubdate3[0]
            $publishr = $pubdate3[1,2,3]
        }

        if($null -eq $publisher){
        $publisher = $publishr
        }
    
    ############################
    # CUE NFO ImageHR Filenames
    ############################
    $fileCUE = "$artist - $album.cue"
    $fileNFO = "$artist - $album.nfo"    
    $imageHR = "$artist - $album.HR.jpg"
 
    ##############################
    # CUE Chapters + new CUE file
    ##############################
    $patternChapter = 'Chapter\s\d+'
    $cueFileTemp = Get-ChildItem -LiteralPath $folder -recurse | Where-Object {$_.name -match "\.(cue)$"}
    If($cueFileTemp){
        $cueContent = Get-Content -Path $cueFileTemp.FullName 
        $chapterCount = ($cueContent | Select-String -Pattern $patternChapter).count
        if($chapterCount -lt '1'){$chapterCount = ''}        
        #write-host $cueContent.count
        $line = $cueContent | Select-String File | Select-Object -ExpandProperty Line
        if($line){
            Remove-Item "$folder\$cueFileTemp" -Force
            $cueContent -replace $line, "FILE `"$fileShort`" WAVE" | Set-Content "$folder\$fileCUE"
        }
        #Write-Host $line
    }

    ###################################
    # Series and Title from album name
    ###################################
    If($album -match 'BK '){
        $series1,$title = $album -split " - "
        #$series2,$seriesNum = $series1 -split ", "
        #$seriesNum = $seriesNum.Replace("BK 0","#").Replace("BK 1","#1")
        #$series = "$series2 $seriesNum"
        $title = $title.Trim()
    }
    Else{$title = $album}

    ###################################
    # Narrator + Reader + Desc info
    ###################################
    If (!($narrator)){write-host "Narrator is Missing " -ForegroundColor Yellow}

    If (($FileExists -eq $False)-and($narrator)-and(!($extension -eq '.aac')))
        {
        write-host "Narrator= " $narrator
        $comments | Out-File -LiteralPath "$folder\desc.txt" -Force -Encoding utf8
        $narrator | Out-File -LiteralPath "$folder\reader.txt" -Force -Encoding utf8
        }
 
    ##################
    # Cover Art files
    ##################
    If (($FolderJPGExists -eq $True)-and($CoverJPGExists -eq $False)) {Copy-Item -literalpath "$folder\folder.jpg" "$folder\cover.jpg"}
    If ($FolderJPGExists -eq $False){Write-Host "Cover art is missing" -ForegroundColor red}
  
    #####################
    # Write out NFO file 
    #####################
    #If($fileNFOExist -eq $False){
    If(($fileNFOExist -eq $False)-or($fileNFOExist -eq $True)){ #overwritting if it NFO exists or not
        Write-Host "*** NFO is being created ***`n" -ForegroundColor Darkyellow

        # Get duration length
        Get-AudioDetails $folder | Sort-Object Duration -Descending | Group-Object Item | ForEach-Object{
            [INT]$a = ($_.Group | Measure-Object Duration -Sum).Sum
            $duration1 = (New-TimeSpan -Minutes $a).ToString()
            $t = $duration1.Split(":")
            $hrs = $t[0]
            $min = $t[1]
            #check to see if hours is greater than 24hrs, if greater, time gets resolved to 1.n, need to convert to 24hour clock format
            $shrs = $hrs.Split(".")
            $splitHr = $shrs[0]
            $splitHr2 = $shrs[1]
            if ($null -ne $splitHr2){
            $newhrs = 24 + $splitHr2
            $duration2 = "$newhrs hours $min minutes"
            }
            elseif($hrs -eq '00'){
                $duration2 = "$min minutes"}
            else{
                $duration2 = "$hrs hours $min minutes"}
        }

        $regexCdate1 = 'c20\d\d'
        $regexCdate2 = 'c19\d\d'
        if($comments -match $regexCdate1){$comments = $comments -replace('c20','(C)20')}
        if($comments -match $regexCdate2){$comments = $comments -replace('c19','(C)19')}
    
$NFOArray = @”
General Information
===================
 Title:                  $title
 Author:                 $albumartists
 Series:                 $series
 Series Position:        $seriesNum
 Read By:                $narrator
 Copyright:              $year
 Audiobook Copyright:    $copyDate    
 Genre:                  $jGenre2
 Publisher:              $publisherTone
 Duration:               $duration2
 Subtitle:               $subtitle
 Chapters:               $chapterCount
 ASIN:                   $asin
 ISBN:                   $isbn

Media Information
=================
 Lossless Encode:        No
 Encoded Codec:          $format $Format_Ver $Format_Prof
 Encoded Library:        $encoded_Lib
 Encoded Bitrate:        $ovAllBitRate kb/s
 Encoded Sample Rate:    $samplingRate kHz
 Encoded Channels:       $channel  $channelDescTone
 Bitrate Mode:           $bitRateMode
 Encoding Tool:          $encodingTool

Book Description
================
$comments
“@

    Write-Host "NFO path $folder\$fileNFO" -ForegroundColor yellow   
    $NFOArray | Out-File -encoding utf8 -LiteralPath "$folder\$fileNFO" -Force

    $comments | Out-File -LiteralPath "$folder\desc.txt" -Force -Encoding utf8
    $narrator | Out-File -LiteralPath "$folder\reader.txt" -Force -Encoding utf8   

} # Write out NFO file

    ##########
    #$imageHD
    ##########
    If (!($imageHRExits)-and($FolderJPGExists)){
    $Image = [System.Drawing.Image]::FromFile("$folder\folder.jpg")
    $ImageWidth = $Image.Width
    $Image.Dispose()
        If ($ImageWidth -gt 800){
         Copy-Item -literalpath "$folder\folder.jpg" "$folder\$imageHR" -Force
         #Write-Host "HR image created $imageHR" -ForegroundColor green
         }
    }
   $imageHR = ""

    } #Ends foreach book
} #Ends foreach folder

Pop-Location
