# my_VbaDrive
Function for checking, if a drive exists and is available.
Find available driveletter

# Examples
Drive C is a harddrive.
Debug.Print drive_Exists("C") ' --> True

Drive H is a disconnected networkdrive
Debug.Print drive_Exists("H") ' --> False
Debug.Print drive_Share_Exists("H") ' --> True
Debug.Print drive_Available("H") ' --> False

Find last available driveletter start with letter U go backword until F.
Debug.Print drive_Last_Letter("U", "F")


