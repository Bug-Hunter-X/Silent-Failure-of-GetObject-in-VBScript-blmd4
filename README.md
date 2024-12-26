# Silent Failure of GetObject in VBScript

This repository demonstrates a subtle error that can occur when using the `GetObject` function in VBScript.  If the specified key doesn't exist, `GetObject` might not raise an error, returning a null value without indication. This can lead to unexpected behavior in your script.

The `bug.vbs` file shows the problematic code. The `bugSolution.vbs` file provides a more robust solution.