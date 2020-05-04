set shell = createobject("wscript.shell")
strtext = inputbox("Spam Message")
strtimes = inputbox("# of msgs")
if not isnumeric(strtimes) then
wscript.quit
end if
msgbox "3 seconds until spam starts."
wscript.sleep (3000)
for i=1 to strtimes
shell.sendkeys(strtext & "{enter}")
wscript.sleep (1)
next