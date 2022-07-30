function Encrypt-Ceasar {
    [OutputType([String])]
    Param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [string]$PlainText,

        [ValidateRange(-25, 25)]
        [int]$Shift = 1
    )

    $Encoded = $PlainText.ToCharArray() | ForEach-Object {
        [int][char]$_ + $Shift
    }        
    
    $Result = -join ( $Encoded | ForEach-Object {
            [char][int]$_ 
        } )

    return $Result
}