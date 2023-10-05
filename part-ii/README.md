# COM Interoperability

### Communication between Visual Basic and C# with events

In this tutorial, we'll delve into the fascinating world of COM Interoperability, demonstrating seamless communication between Visual Basic 6 (VB6) and C# applications. 
Throughout the tutorial, you'll learn the crucial steps to establish this communication channel, unlocking the potential for VB6 and C# applications to work in harmony. You'll witness the power of COM Interoperability as the C# library displays the received message in a message box, showcasing a practical integration between legacy and modern programming environments.

> Check the repository for the full source code 
>
> https://github.com/nikosportolos/COM-Interoperability-in-.NET/tree/main/part-ii


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

1. Create your Event interface

```csharp
[Guid("123456789-1234-1234-1234-123456789123")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IVB6InteropEvents
{
    [DispId(1)]
    void SampleEvent(string Message);
}
```

2. Implement Events 

   - Create delegate
     
     ```csharp
     public delegate void SampleEventHandler(string Message);
     ```

   - Create event

      ```csharp
      public new event SampleEventHandler SampleEvent = null;
      ```

    - Create event trigger

      ```csharp
      public void FireSampleEvent(string Message);

      [Guid("7E6D6368-0033-49F6-9FE3-B2D409572869")]
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
                //MessageBox.Show(Message);
                FireSampleEvent("I received the message: " + Message);
            }
            catch (Exception ex)
            {
                throw new Exception("Exception occured in SampleMethod(): ", ex);
            }
        }
        
        // Create delegate
        [ComVisible(true)]
        public delegate void SampleEventHandler(string Message);
        
        // Create event
        public new event SampleEventHandler SampleEvent = null;
        
        // Create event trigger
        public void FireSampleEvent(string Message)
        {
            try
            {
                if (SampleEvent != null)
                    SampleEvent(Message);
            }
            catch (Exception ex)
            {
                throw new Exception("Exception occured in FireSampleEvent(): ", ex);
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

```vb
Option Explicit

' Declare a VB6Interop object
Public WithEvents VB6 As VB6Interop.VB6Interop
 
Private Sub Command1_Click()
    ' Use SampleMethod() of VB6Interop
    VB6.SampleMethod "Hello VB6!"
End Sub
 
Private Sub Form_Load()
    ' Initialize VB6Interop object
    Set VB6 = New VB6Interop.VB6Interop
End Sub
 
' Implement SampleEvent
Private Sub VB6_SampleEvent(ByVal Message As String)
    MsgBox Message
End Sub
```


## Demo

<a href="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/vb_05.png" target="_blank">
    <img src="https://raw.githubusercontent.com/nikosportolos/COM-Interoperability-in-.NET/main/Resources/screenshots/vb_05.png" width="750" alt="vb_05">
</a> 
