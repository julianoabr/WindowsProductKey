<#
.Synopsis
   Get Windows Product Key Local or Remote
.DESCRIPTION
   Get Windows Product Key Local or Remote
.EXAMPLE
   Run The Script. Type Local or Remote and Get Product KEY
.EXAMPLE
   Another example of how to use this cmdlet
.AUTHOR
   Juliano Alves de Brito Ribeiro (Find me at: julianoalvesbr@live.com or https://github.com/julianoabr)
.VERSION
 0.2
.ENVIRONMENT
 PROD

.TO THINK

This is was written more than 2000 years ago, and we are so close to this. Revelation Chapter 13. 

15 The second beast was given power to give breath to the image of the first beast, so that the image could speak and cause all who refused to worship the image to be killed.
16 It also forced all people, great and small, rich and poor, free and slave, to receive a mark on their right hands or on their foreheads, 
17 so that they could not buy or sell unless they had the mark, which is the name of the beast or the number of its name.

#>


#GET OS VERSION
$rWmiOSVersion = {param([string]$rServer=$env:COMPUTERNAME) 
try { $rOsVersion = ((Get-WmiObject -ComputerName $rServer -Class win32_operatingsystem -ErrorAction Stop).Version).ToString()}
catch{$rOsVersion = "Failure" }
return $rOsVersion
}


function PSPing([string]$hostname, [int]$timeout = 50) {
$ping = New-Object System.Net.NetworkInformation.Ping #creates a ping object
try { $pingResult = $ping.send($hostname, $timeout).Status.ToString() }
catch { $pingResult = "Failure" }
return $pingResult
}


function Get-WindowsProductKey
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias("CN", "Computer")]
        [System.String]
        $rServerName = $env:COMPUTERNAME,

        [Parameter(Position = 1, ValueFromPipeline = $true)]
        [System.String]
        $rOsVersionInfo

    )  
 

  # implement decoder
  $code = @'
// original implementation: https://github.com/mrpeardotnet/WinProdKeyFinder
using System;
using System.Collections;

  public static class Decoder
  {
        public static string DecodeProductKeyWin7(byte[] digitalProductId)
        {
            const int keyStartIndex = 52;
            const int keyEndIndex = keyStartIndex + 15;
            var digits = new[]
            {
                'B', 'C', 'D', 'F', 'G', 'H', 'J', 'K', 'M', 'P', 'Q', 'R',
                'T', 'V', 'W', 'X', 'Y', '2', '3', '4', '6', '7', '8', '9',
            };
            const int decodeLength = 29;
            const int decodeStringLength = 15;
            var decodedChars = new char[decodeLength];
            var hexPid = new ArrayList();
            for (var i = keyStartIndex; i <= keyEndIndex; i++)
            {
                hexPid.Add(digitalProductId[i]);
            }
            for (var i = decodeLength - 1; i >= 0; i--)
            {
                // Every sixth char is a separator.
                if ((i + 1) % 6 == 0)
                {
                    decodedChars[i] = '-';
                }
                else
                {
                    // Do the actual decoding.
                    var digitMapIndex = 0;
                    for (var j = decodeStringLength - 1; j >= 0; j--)
                    {
                        var byteValue = (digitMapIndex << 8) | (byte)hexPid[j];
                        hexPid[j] = (byte)(byteValue / 24);
                        digitMapIndex = byteValue % 24;
                        decodedChars[i] = digits[digitMapIndex];
                    }
                }
            }
            return new string(decodedChars);
        }

        public static string DecodeProductKey(byte[] digitalProductId)
        {
            var key = String.Empty;
            const int keyOffset = 52;
            var isWin8 = (byte)((digitalProductId[66] / 6) & 1);
            digitalProductId[66] = (byte)((digitalProductId[66] & 0xf7) | (isWin8 & 2) * 4);

            const string digits = "BCDFGHJKMPQRTVWXY2346789";
            var last = 0;
            for (var i = 24; i >= 0; i--)
            {
                var current = 0;
                for (var j = 14; j >= 0; j--)
                {
                    current = current*256;
                    current = digitalProductId[j + keyOffset] + current;
                    digitalProductId[j + keyOffset] = (byte)(current/24);
                    current = current%24;
                    last = current;
                }
                key = digits[current] + key;
            }

            var keypart1 = key.Substring(1, last);
            var keypart2 = key.Substring(last + 1, key.Length - (last + 1));
            key = keypart1 + "N" + keypart2;

            for (var i = 5; i < key.Length; i += 6)
            {
                key = key.Insert(i, "-");
            }

            return key;
        }
   }
'@
  # compile C#
  Add-Type -TypeDefinition $code
 
  # get raw product key
  if ($rServerName -eq $env:COMPUTERNAME)
  {
      
       $digitalId = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name DigitalProductId).DigitalProductId

  }#If Get Local Product Key
  else{
  
       $digitalId = Invoke-Command -ComputerName $rServerName -ScriptBlock{(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name DigitalProductId).DigitalProductId}
  
  }#Else Get Remote Product Key
 
  
  if (($rOsVersionInfo -match '^[5-6]\.[0-2].\.*') -or ($rOsVersionInfo -match '^6\.[0-1].\.*'))
  {
    # use static C# method
    [Decoder]::DecodeProductKeyWin7($digitalId)
  }
  elseif(($rOsVersionInfo -match '^6\.[2-3].\.*') -or ($rOsVersionInfo -match '^10.0.\.*'))
  {
    # use static C# method:
    [Decoder]::DecodeProductKey($digitalId)
       
  }
  else{
  
      Write-Host "I can't decote the Product ID of Server: $rServerName" -ForegroundColor White -BackgroundColor Red
  
  }

}#End of Function



#RUN SCRIPT LOCAL OR REMOTE
do
{
 
 $answerLR = ""
 
 $tmpAnswerLR = ""

 [System.String]$tmpanswerLR = Read-Host "Type [LOCAL] to get Product Key of this Computer. Type [REMOTE] to get Product Key of Remote Computer"
 
 $answerLR = $tmpanswerLR.ToUpper() 
  
}
while ($answerLR -notmatch '^(?:LOCAL\b|REMOTE\b)')


if ($answerLR -match 'LOCAL'){

    $localOSVersion = Invoke-Command -ScriptBlock $rWmiOSVersion

    Get-WindowsProductKey -rOsVersionInfo $localOSVersion

}

if ($answerLR -match 'REMOTE'){

    $rServer = Read-Host "Type the Remote Server Name to Get Product ID"

    $result = PSPing -hostname $rServer -timeout 10

    if ($result -eq 'SUCCESS')
    {
        
        #Get Remote OS Version
        $rJob = Start-Job -ScriptBlock $rWmiOSVersion -ArgumentList $rServer

        $rOsVInfo = Wait-Job $rJob -Timeout 15 | Receive-Job

        if ($null -ne $rOsVInfo){
        
            Get-WindowsProductKey -rServerName $rServer -rOsVersionInfo $rOsVInfo
        
        
        }#Get Product Key if Result Not Null
        else{
        
             Write-Host "I can't connect to server: $rServer. Please verify WMI" -ForegroundColor White -BackgroundColor Red
            
        
        }#Say that cannot connect to remote server

        $rJobID = $rJob.Id

        Get-Job -Id $rJobID | Remove-Job -Force


    }#if ping
    else{
    
        Write-Host "I can't ping the server: $rServer" -ForegroundColor White -BackgroundColor Red
    
    }#else ping
    

}#end of test remote
