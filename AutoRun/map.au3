Opt("TrayIconHide", 1)          ;0 - отображать, 1 - скрыть

If StringUpper(@ComputerName) <> "UBYTKI_CHCH" Then

	DriveMapAdd ( "i:", "\\172.16.1.2\doc" ,1);
	DriveMapAdd ( "z:", "\\172.16.1.2\ingos" ,1);

EndIf


