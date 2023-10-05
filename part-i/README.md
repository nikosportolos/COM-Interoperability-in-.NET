# COM Interoperability

### Basic communication between Visual Basic and C#

This tutorial provides a quick and easy guide for basic communication between Visual Basic 
and C# utilizing the COM Interoperability in .NET.

We'll focus on building a C# class library designed to receive messages from a VB6 application. 
The VB6 component features a straightforward form housing a button that dispatches a message to the C# library.

> Check the repository for the full source code 
>
> https://github.com/nikosportolos/COM-Interoperability-in-.NET/tree/main/part-i

## Table of Contents

- [.NET class library](#net-class-library)
  - [Setup .NET project](#setup-net-project)
  - [C# source code](#c-source-code)
- [VB6 winform app](#vb6-winform-app)
  - [Setup VB6 project](#setup-vb6-project)
  - [VB6 winform app](#vb6-winform-app)
- [Demo](#demo)

---

## .NET class library

### Setup .NET project

1. Create a new C# Project in Visual Studio

    <a href="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/dotnet_01.png" target="_blank">
    <img src="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/dotnet_01.png" width="750" alt="dotnet_01">
    </a>


2. Select **Class Library**

    <a href="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/dotnet_02.png" target="_blank">
    <img src="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/dotnet_02.png" width="750" alt="dotnet_02">
    </a> 


3. Make assembly **COM Visible**

  - Right-Click on Project
  - Properties (Alt+Enter)

    <a href="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/dotnet_03.png" target="_blank">
    <img src="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/dotnet_03.png" width="750" alt="dotnet_03">
    </a>

    - Application tab
    - Assembly Information button

    <a href="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/dotnet_04.png" target="_blank">
    <img src="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/dotnet_04.png" width="750" alt="dotnet_04">
    </a> 
    
    - Check option *"Make assembly COM Visible"*

    <a href="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/dotnet_05.png" target="_blank">
    <img src="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/dotnet_05.png" width="750" alt="dotnet_05">
    </a> 


4. Register for COM Interop

    - Build tab
    - Output section
    - Check option *"Register for COM interop"*

    <a href="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/dotnet_06.png" target="_blank">
    <img src="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/dotnet_06.png" width="750" alt="dotnet_06">
    </a> 


### C# source code

1. Access **InteropServices** Namespace

```csharp
using System.Runtime.InteropServices
```

2. Create your Class interface

```csharp
[Guid("123456789-1234-1234-1234-123456789123")]
[ComVisible(true)]
public interface IVB6Interop
{
    [DispId(1)]
    void SampleMethod(string Message);
}
```

3. Implement the **IVB6Interop** interface


```csharp
[Guid("123456789-1234-1234-1234-123456789123")]
[ProgId("VB6Interop")]
[ClassInterface(ClassInterfaceType.None)]
[ComSourceInterfaces(typeof(IVB6InteropEvents))]
[ComVisible(true)]
public class VB6Interop : IVB6Interop
{
    public void SampleMethod(string Message)
    {
        try
        {
            MessageBox.Show(Message);
        }
        catch (Exception ex)
        {
            throw new Exception("Exception occured in SampleMethod(): ", ex);
        }
    }
}
```

---

## VB6 winform app

### Setup VB6 project

1. Create a new VB6 Project

  - Select Standard EXE
  <a href="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/vb_01.png" target="_blank">
    <img src="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/vb_01.png" width="750" alt="vb_01">
  </a> 

- Add *VB6Interop.dll* reference
  - Project
  - References

    <a href="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/vb_02.png" target="_blank">
      <img src="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/vb_02.png" width="750" alt="vb_02">
    </a> 

  - Check **VB6Interop** Reference

    <a href="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/vb_03.png" target="_blank">
      <img src="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/vb_03.png" width="750" alt="vb_03">
    </a> 


### VB6 source code

1. Declare a **VB6Interop** object

```vb
Public VB6 As VB6Interop.VB6Interop
```

2. Initialize VB6Interop object

```vb
Private Sub Form_Load()
    Set VB6 = New VB6Interop.VB6Interop
End Sub
```

3. Use SampleMethod() of VB6Interop on button click

```vb
Private Sub Command1_Click()
    VB6.SampleMethod "Hello VB6!"
End Sub
```

*…and we’re done!*


## Demo

<a href="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/vb_04.png" target="_blank">
    <img src="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/vb_04.png" width="750" alt="vb_04">
</a> 

