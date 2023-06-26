# my_VbaDrive
Function for checking, if a drive exists and is available.
Find available driveletter

# Examples
Drive C is a harddrive.<br>
Debug.Print drive_Exists("C") ' --> True<br>
<br>
Drive H is a disconnected networkdrive<br>
Debug.Print drive_Exists("H") ' --> False<br>
Debug.Print drive_Share_Exists("H") ' --> True<br>
Debug.Print drive_Available("H") ' --> False<br>
<br>
Find last available driveletter start with letter U go backword until F.<br>
Debug.Print drive_Last_Letter("U", "F")<br>


