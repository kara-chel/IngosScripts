Opt("TrayIconHide", 1)          ;0 - ����������, 1 - ������

#include <Process.au3>

_RunDOS('\\172.16.1.2\Doc\Scripts\Bginfo.exe \\172.16.1.2\Doc\Scripts\Bginfo.bgi /NOLICPROMPT /TIMER:0')