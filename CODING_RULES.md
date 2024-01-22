## Intention
It's the result of a long time experience: Rules ease code reading and understanding even after years. 
Exceptions are the implication of rules (otherwise they would be called laws). Every now and then there might be a good reason not to follow a rule dogmatic.
## Avoiding defaults
| What | Why |
|------|---------|
|Option&nbsp;Explicit| Enforcing declaration of variables avoids default Variant declaration|
|Public or Private| Avoiding the default Public when a procedure is used only within the component |
|ByVal or ByRef| The explicit specification of a read-only argument versus a returned-value argument clearly shows the intention. |

## Upper/Lower case rules for declarations
| What | Reason, details |
|------|---------|
|UPPER_CASE_CONSTANTS| Clearly separate them from variables, underscores support readability|
|lower_case_arguments| Clearly indicates them when used in code lines, underscores support readability, a prefix avoids any mismatch with a system argument with the same name thereby changing the upper/lower case writing of them |
| ProcedureNames | Upper an lower case letters support readability, no underscores clearly identifies them as Sub, Function, or Property|
| lVariable As Long | The lower case prefix identifies the as variables |
| A three letter prefix for variables declaring an object| rng Range<br>wbk Workbook<br>wsh Worksheet<br>nme Name|

## Other
| What | Reason, details |
|------|---------|
Named arguments| When a Sub, Function, or Property has more than one argument, named arguments (`argument:=xxxxx`) makes the code independent from the arguments position which may change when modified.|
|Object existence checks| Preferably in dedicated functions, not only return True or False but also the existing object as a ByRef argument.<br>
| Worksheet identification | Exclusively by their CodeName. Makes it independent from the position and from the Name, both may be changed by the user|
|Error Handling| An elaborated error handling supports the path to the error, the line where the error occured and regression testing by not displaying asserted errors|
|Execution Trace| A built-in execution trace, activated via a Conditional Compile Argument supports performance issues|
|Common Components| Developed, supplemented, and carefully regression tested, they became a true efficiency boost for VB-Project TS development.|


