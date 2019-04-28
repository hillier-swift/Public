#usage Ping-range.ps1 -subnet 10.14.0
param
([parameter(Mandatory=$true)][string]$subnet,
[string]$outfile = "C:\Stuff\PingRange.csv"
)

$results = 1..254 | foreach { new-object psobject -prop @{Address=”$subnet.$_”;Ping=(test-connection “$subnet.$_” -quiet -count 1)}}

$results | Export-csv $outfile