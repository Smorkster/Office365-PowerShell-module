#Persistent
#SingleInstance force
Loop
{
	WinWaitActive, ahk_class WindowsForms10.Window.8.app.0.34f5582_r12_ad1
	{
		Sleep 2000
		WinGetTitle, T, A
		test := "Logga in på ditt konto"
		if (T = test)
		{
			Send {Alt down}{Ctrl down}a
			;Sleep 100
			Send {Alt up}
			Send {Ctrl up}
		}
		WinWaitClose, Logga in på ditt konto
	}
}
