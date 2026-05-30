# 🕶️ vba-pw3270-utils // THE HLLAPI INJECTOR

[![Status](https://shields.io)]()
[![Engine](https://shields.io)]()
[![Interface](https://shields.io)]()
[![Target](https://shields.io)]()

> "Forget slow abstraction layers. Bind the core memory strings directly. We are inside the mainframe's veins."

Welcome to **vba-pw3270-utils**. This architecture contains customized Excel VBA routines optimized for direct memory interaction with the `pw3270` IBM terminal emulator via native Windows `libhllapi32.dll` function declarations.

---

## ⚡ CORE ARCHITECTURE

* **Source Engine:** Microsoft Excel VBA (Class Modules)
* **Low-Level Interface:** HLLAPI (High-Level Language Application Programming Interface)
* **Subsystem Library:** `libhllapi32.dll` (or `libhllapi.dll` depending on target installation)

---

## 💾 CYBER-INFRASTRUCTURE FILES

* `/src/clsLibHllapi.cls` — The core binding class layer mapping raw terminal hex memory addresses into clean VBA objects.
* `/src/Main_Payloads.bas` — Standard procedural modules holding automation loops, bulk execution commands, and form-submitting tasks.

---

## 🛠️ JACKING INTO THE MATRIX

1. **Clone the Source Matrix:**
   ```bash
   git clone https://github.com
   ```
2. **Access the Subsystem:** Boot Excel ➔ Press `ALT + F11` to deploy the VBA IDE.
3. **Inject the Class Module:** Right-click the project tree ➔ *Import File* ➔ Choose `clsLibHllapi.cls`.
4. **Target Alignment Check:** If your local machine environment uses the 64-bit/alternative compile path, ensure the `Lib` directives point to `libhllapi.dll` instead of `libhllapi32.dll`.

---

## 💻 CODE PAYLOAD (INITIALIZING BUFFER STREAM)

```vba
' // DIRECT MEMORY INJECTION METHOD
Sub ExecuteMainframeBridge()
    Dim Terminal As New clsLibHllapi
    
    ' Establish native link to terminal window session
    If Terminal.Connect("pw3270:A") Then
        ' Wait until system responds
        Terminal.WaitHostOK 5
        
        ' Inject data stream at coordinates: Row 10, Column 5
        Terminal.PutString 10, 5, "ACCESS_GRANTED"
        Terminal.Enter
        
        Debug.Print "[SUCCESS] Memory handshake complete. Buffer active."
    Else
        Debug.Print "[CRITICAL ERROR] Handshake timed out. Session not targeted."
    End If
End Sub
```

---

## 🚨 SYSTEM WARNING

You are communicating directly with underlying system terminal screen buffers. Executing massive unthrottled loops can trigger the host firewall or data flooding protection systems. Always execute tests inside a secure staging environment.

---

## 📡 TERMINAL TRANSMISSION

* Architecture designed & adapted from original logic by **Erick Lopes de Souza**
* Maintained by [@YOUR-GITHUB-HANDLE](https://github.com/fnusda)
* Project Status: Operating Smoothly / Active
