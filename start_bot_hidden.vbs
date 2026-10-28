' Starts start_bot.cmd with NO console window at all.
' A Task Scheduler task that runs the .cmd directly opens a visible console
' in the user session, and closing that window kills the bot. WScript.Shell
' with window style 0 leaves nothing to close, while the bot still runs in
' the interactive session (which the 1C client needs - session 0 has no
' desktop, see start_bot.cmd).
Dim fso, here, sh
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
Set sh = CreateObject("WScript.Shell")
sh.Run """" & here & "\start_bot.cmd""", 0, False
