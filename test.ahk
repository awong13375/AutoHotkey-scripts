;Extract username and password from text file
Contents := FileOpen("Settings.txt", "r")
A := Contents.Readline()
B := Contents.Readline()

A := Substr(A, 10)
B := Substr(B, 10)

A := Trim(A)
B := Trim(B)

A := Trim(A, "'")
B := Trim(B, "'")


