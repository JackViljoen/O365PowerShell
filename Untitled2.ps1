function Check-Value {
    Param([parameter(position = 0)] $V)
     #does some checkifgn on the type of a value its a sort of a convenience function  
    if ($null -ne $V){
        $T = $V.GetType();
        Switch ($T.Name)
            {
                "Int32"    {
                              if($V -eq 0){$R = "Zero Integer"} else {$R = "Integer"}
                           }
                "Double"   {$R = "Number"}
                "Decimal"  {$R = "Number"}
                "String"   {
                                $N = 0
                                $result = [System.Int32]::TryParse($V,[ref]$N)
                                if($result){
                                    if($N -eq 0){$R = "Zero Integer"} else {$R = "Integer"}
                                } else {
                                    $result = [System.Decimal]::TryParse($V,[ref]$N)
                                    if($result){
                                        if($N -eq 0){$R = "Zero Number"} else {$R = "Number"}
                                    } else {
                                        if($V.length -eq 0){$R = "Empty String"} else {$R = "String"}
                                    }
                                }
                            }
                "DateTime"  {$R = "Date"}

                default     {$R = "Other"}
            }
    } else {
       $R = "null"
    }
    write-host $R $V -ForegroundColor Yellow
    return $R 
}

check-value ""

check-value 0
check-value "0"

check-value 00.001

check-value 00.00

check-value "34.55"


check-value 12


check-value "34.45667"


check-value "23F56"


check-value "this is a tim thing string"