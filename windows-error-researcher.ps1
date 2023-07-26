# Clear the console screen for a fresh start
Clear-Host

function Get-SystemEvents {
    Get-WinEvent -LogName System -FilterXPath "*[System[Level=2 or Level=3]]" |

    ForEach-Object {
        [PSCustomObject]@{
            EventID = $_.Id
            Message = $_.Message
            Count = 1
        }
    } |

    Group-Object -Property EventID |

    Sort-Object -Property Count -Descending |

    Select-Object @{Name='Count'; Expression={$_.Count}},
    @{Name='Event ID'; Expression={$_.Group[0].EventID}},
    @{Name='Message'; Expression={$_.Group[0].Message}} -First 5
}

Write-Host "Checking your computer for most frequent errors..." -ForegroundColor Yellow

# Pause so user can read
Start-Sleep 3

# Without Out-Host piped, Write-Host will steal the spotlight and hide function output
Get-SystemEvents | Out-Host

Write-Host "Please wait. I will retrieve your solutions..." -ForegroundColor Green

# Put the events into a variable to iterate through
$Events = Get-SystemEvents

# Pause so user can read
Start-Sleep 5

foreach ($Event in $Events) {
    $SearchString = "$($Event.EventID) $($Event.Message)"
    $SearchString = $SearchString -replace ' ', '+'
    
    # Open the search results in Microsoft Edge (in case default browser is set to a malware browser)
    Start-Process "msedge.exe" -ArgumentList "https://www.bing.com/search?q=$SearchString+site:microsoft.com"
}
