param(                                    #Модуль param обязатьно должен идти первой строкой, иначе работь не будет
 [Parameter(Mandatory=$true)]
    [String]$input_path,
 [Parameter(Mandatory=$true)]
    [String]$output_path
)
function Get-cam_obj{
    param([Parameter(Mandatory=$true)]$user)
        $cam='CAMID("MAIN:u:'
        $a=[string]$user.ObjectGUID
        $a=$a.split('-')
        $a[0]=$a[0][6..7]+$a[0][4..5]+$a[0][2..3]+$a[0][0..1]
        $a[1]=$a[1][2..3]+$a[1][0..1]
        $a[2]=$a[2][2..3]+$a[2][0..1]
        $b=$cam+[string]$a[0].replace(' ','')+[string]$a[1].replace(' ','')+[string]$a[2].replace(' ','')+[string]$a[3].replace(' ','')+[string]$a[4].replace(' ','')+'")'
        $ADName='AD/'+[string]$obj.Name
        $user | Add-Member -MemberType NoteProperty -Name CAMID -Value $b -Force
        $user | Add-Member -MemberType NoteProperty -Name ADName -Value $ADName -force
        return $user

}
function Get-cam{
    param(
    [Parameter(Mandatory=$true)]
      [String]
      $oblectguid)
      $a=[string]$objectguid
        $a=$a.split('-')
        $a[0]=$a[0][6..7]+$a[0][4..5]+$a[0][2..3]+$a[0][0..1]
        $a[1]=$a[1][2..3]+$a[1][0..1]
        $a[2]=$a[2][2..3]+$a[2][0..1]
        $b=$cam+[string]$a[0].replace(' ','')+[string]$a[1].replace(' ','')+[string]$a[2].replace(' ','')+[string]$a[3].replace(' ','')+[string]$a[4].replace(' ','')+'")'
        return $b
        }
function Get-obj_guid{
    param(
    [Parameter(Mandatory=$true)]
      [String]
      $camid)
        $a=([string]$camid[14..21]).replace(' ',''),([string]$camid[22..25]).replace(' ',''),([string]$camid[26..29]).replace(' ',''),([string]$camid[30..33]).replace(' ',''),([string]$camid[34..45]).replace(' ','')
        $a[0]=$a[0][6..7]+$a[0][4..5]+$a[0][2..3]+$a[0][0..1]
        $a[1]=$a[1][2..3]+$a[1][0..1]
        $a[2]=$a[2][2..3]+$a[2][0..1]
        $itog=([string]$a[0]).replace(' ','')+'-'+([string]$a[1]).replace(' ','')+'-'+([string]$a[2]).replace(' ','')+'-'+[string]$a[3]+'-'+[string]$a[4]
        return $itog
      }

Import-Module activedirectory
Remove-item "$output_path"
if(Test-Path -Path $input_path){
    $file=Import-Csv -Path $input_path -Delimiter "," -Encoding default  #соблюдение системных кодировок
    $id=$file[0].psobject.properties.name[0] #Убрать решетку если больше 1 столбца
        if($id -eq 'mail'){
        foreach($row in $file){
            $email=[string]$row."$id"
            $obj=Get-AdUser -Filter { mail -Like $email}  -Properties mail, ObjectGUID , Name ,title, Enabled 
            Get-cam_obj -user $obj| select CAMID,mail,ADName,ObjectGUID,Name,Title,Enabled | Export-csv -path $output_path -Append -Encoding Default -NoTypeInformation
        }
        }elseif($id -eq 'objectguid'){
        foreach($row in $file){
            $objectguid=[string]$row."id"
            $obj=Get-AdUser -identity $objectguid -Properties mail, ObjectGUID , Name ,title, Enabled
            Get-cam_obj -user $obj| select CAMID,mail,ADName,ObjectGUID,Name,Title,Enabled | Export-csv -path $output_path -Append -Encoding Default -NoTypeInformation
            }
        
        }elseif($id -eq 'CAMID'){
            foreach($row in $file){
            try{
            $camid=[string]$row."$id"
            $objectguid=Get-obj_guid -camid $camid
            $obj=Get-AdUser -identity $objectguid -Properties mail, ObjectGUID , Name ,title, Enabled
            Get-cam_obj -user $obj| select CAMID,mail,ADName,ObjectGUID,Name,Title,Enabled | Export-csv -path $output_path -Append -Encoding Default -NoTypeInformation
            }catch{
                   try{
                   $email=[string]$row."Mail"
                   $obj=Get-AdUser -Filter { mail -Like $email}  -Properties mail, ObjectGUID , Name ,title, Enabled
                   Get-cam_obj -user $obj| select CAMID,mail,ADName,ObjectGUID,Name,Title,Enabled | Export-csv -path $output_path -Append -Encoding Default -NoTypeInformation


                  }catch{
                  $row| Add-Member -MemberType NoteProperty -Name ObjectGUID -Value 'removed from ad' -Force
                  $row| Add-Member -MemberType NoteProperty -Name Name -Value 'removed from ad' -Force
                  $row| Add-Member -MemberType NoteProperty -Name Title -Value 'removed from ad' -Force
                  $row| Add-Member -MemberType NoteProperty -Name Enabled -Value 'removed from ad' -Force
                  $row|Export-csv -path $output_path -Append -Encoding Default -NoTypeInformation
            }
            
            }
            }
        }else{
        }
}else{
Write-Error 'incorrect input path'
}
'end'
