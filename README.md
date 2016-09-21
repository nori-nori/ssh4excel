This DLL is a COM wrapper of SSH.NET.

On Excel Macro/VBA, to connect a target via SSH.

For example, to connect a router or switch, other net devices.


# Sample(screenshot)

![](https://github.com/nori-nori/ssh4excel/example/ssh.gif)

***


# requirement
[SSH.NET](https://github.com/sshnet/SSH.NET) (Renci.SshNet.dll)

it's SSH library in C#.

SSH4Excel is a COM wrapper for the DLL.

***

#usage
##regist DLL(COM)

it's needed to refer from Excel


cf. regist command on windows10(32bit)

`C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm ssh4excel.DLL /codebase /tlb`

The path depends on your system, 64bit or 32bit, .NETFramework version.

Please check your system.


##code
- ssh open
- expect
- ssh close

it's the same as usual Expect code.


### sample
See VBA code on [ssh.xlsm](https://github.com/nori-nori/ssh4excel/example/ssh.xlsm).

The code is to access Cisco Router.



