# my_VbaDrive
Function for checking, if a drive exists and is available.
Find available driveletter

# Examples
<b>Drive C is a harddrive.<br></b>
Debug.Print drive_Exists("C") ' --> True<br>
<br>
<b>Drive H is a disconnected networkdrive<br></b>
Debug.Print drive_Exists("H") ' --> False<br>
Debug.Print drive_Share_Exists("H") ' --> True<br>
Debug.Print drive_Available("H") ' --> False<br>
<br>
<b>Find last available driveletter start with letter U go backward until F.<br></b>
Debug.Print drive_Last_Letter("U", "F")<br>


