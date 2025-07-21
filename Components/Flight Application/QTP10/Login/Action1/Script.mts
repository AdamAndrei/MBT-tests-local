OpenApp "C:\Program Files (x86)\HP\QuickTest Professional\samples\flight\app\flight4a.exe"
Dialog("Login Dialog").Activate
Dialog("Login Dialog").WinEdit("Agent Name:").Set Parameter("AgentName")
Dialog("Login Dialog").WinEdit("Password:").Set Parameter("AgentPassword")
Dialog("Login Dialog").WinButton("OK").Click
