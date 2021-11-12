Dim fso, ret
Const forWriting = 2
Const forReading = 1
set fso1 = createobject("scripting.filesystemobject")
set ret1 = fso1.OpenTextFile ("C:\Users\ahmed.khalifa\ahmedWs\oddNumber.txt", forReading)
var1= ret1.ReadAll
set fso = createobject("scripting.filesystemobject")
set ret = fso.OpenTextFile ("C:\Users\ahmed.khalifa\ahmedWs\oddNumber.txt", forWriting)
ret.Write var1+1
intresult = (Var1 + 1) mod 2
if intresult = 0 then
reporter.reportevent micPass, "result","The number is even"
else
reporter.reportevent micFail, "result","The number is even"
End if
