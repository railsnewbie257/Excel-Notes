[<h2>Good Patterns For VBA Error Handling</h2>](https://stackoverflow.com/questions/1038006/good-patterns-for-vba-error-handling)


  - On Error Goto ErrorHandlerLabel
  - Resume (Next | ErrorHandlerLabel)
  - On Error Goto 0 (disables current error handler)
  - Err object

The Err object's properties are <b>normally reset to zero or a zero-length string in the error handling routine</b>, but it can also be done explicitly with <b>Err.Clear</b>.

Errors in the error handling routine are terminating.

The range 513-65535 is available for user errors. For custom class errors, you add vbObjectError to the error number.

For not implemented interface members in a derived class, you should use the constant E_NOTIMPL = &H80004001.

<hr>

[<h2>Exception and Error Handling in Visual Basic</h2>](https://msdn.microsoft.com/en-us/library/s6da8809.aspx)

<b>In This Section</b>

Introduction to Exception Handling (Visual Basic)

Summarizes how to handle exceptions in your programs.

Choosing When to Use Structured and Unstructured Exception Handling (Visual Basic)

Describes both types of exception handling and suggests when they are most appropriate.

Types of Errors (Visual Basic)

Provides an overview of syntax errors, run-time errors, and logic errors.
    
Configuring Warnings in Visual Basic

Details how to turn compiler warnings on and off in Visual Basic.

Structured Exception Handling Overview for Visual Basic

Discusses and demonstrates structured exception handling in Visual Basic.

Unstructured Exception Handling Overview (Visual Basic)

Discusses and demonstrates unstructured exception handling in Visual Basic.
    
<b>Related Sections</b>

Generate From Usage

Describes how to generate a stub for an undefined class, constructor, method, property, field, or enum.

Debugger Roadmap

Discusses the fundamentals of using the Visual Studio debugger. Topics include debugging basics, execution control, attaching to a running program, Just-In-Time debugging, launching the debugger automatically, dumps, breakpoints, inspecting your program, handling exceptions, Edit and Continue, and using expressions in the debugger.

Just-In-Time Debugging

Describes just-in-time debugging, a feature that launches the Visual Studio debugger automatically when a program running outside Visual Studio encounters a fatal error.

Debugging Managed Code

Covers common debugging problems and techniques for managed applications.
    
Exception Assistant

Describes the Exception Assistant feature, which facilitates troubleshooting run-time errors.
    
Edit and Continue

Describes Edit and Continue, a time-saving feature that allows you to make changes to source code while the program is in break mode.
    
COM and ActiveX Debugging

Provides tips on debugging COM applications and ActiveX controls.
