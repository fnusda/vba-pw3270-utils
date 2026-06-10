# 🚀 Enterprise PW3270 RPA Framework

## 📖 Overview
This repository contains a highly modular, enterprise-grade Robotic Process Automation (RPA) toolkit built in VBA. It is designed to automate complex data entry, extraction, and financial validation tasks within corporate mainframe terminal systems using the PW3270 emulator and `libhllapi.dll`.

Originally a monolithic script of nearly 2,000 lines, this project has been completely refactored into a scalable, DRY (Don't Repeat Yourself) architecture. It features robust error handling, 64-bit Windows API compatibility, and dynamic UI performance optimization.

## 🏗️ Architecture & Modules
The framework is divided into logical, single-responsibility modules to ensure easy maintenance and scalability:

* **`clsTerminalMacros`**: The core Object-Oriented wrapper for `libhllapi.dll`. Handles 32/64-bit API calls, terminal synchronization, and advanced string-tokenized keystrokes.
* **`Helper_Core`**: Global utilities for managing Excel performance, dynamic row tracking, and UI progress bars.
* **Data Logistics**: Modules (`Data_Preparation`, `Data_Formatting_Chain`, `Data_Management`) dedicated to sanitizing raw reports, executing Regex validations, and archiving processed records.
* **Terminal Automation Engine**: Specialized modules broken down by business logic:
    * `Terminal_Queries` & `Terminal_Identity_Queries`
    * `Terminal_Actions` & `Terminal_Updates`
    * `Terminal_Billing`, `Terminal_Invoicing_And_Retentions`, & `Terminal_Tariff_Simulation`
    * `Terminal_Locations` & `Terminal_Infrastructure`
* **External Integrations**: Modules to handle external processes outside the terminal (`Web_Automation` for Chrome/Edge, `API_Integration` for JSON parsing, `Integration_Office` for Word/Acrobat PDF extraction).

## ⚙️ Prerequisites
To run this framework, your environment must have the following configured:

1.  **PW3270 Terminal Emulator**: Must be installed, with `libhllapi.dll` accessible in the system paths (e.g., `C:\Program Files\PW3270\` or `SysWOW64`).
2.  **Microsoft Excel**: 32-bit or 64-bit (API calls are `PtrSafe`).
3.  **VBA References**: The following object libraries must be enabled in the VBA Editor (`Tools > References`):
    * `Microsoft Scripting Runtime` (for FileSystemObject)
    * `Microsoft Word XX.0 Object Library` (for document parsing)
    * `Acrobat` (for PDF text extraction)
4.  **VBA-JSON**: Requires the standard [VBA-JSON library](https://github.com/VBA-tools/VBA-JSON) to be present in the project for API integrations.

## 🚀 Getting Started
1.  Clone this repository to your local machine.
2.  Import the `.cls` and `.bas` modules into your Excel VBA Project.
3.  Ensure your PW3270 terminal is open and authenticated.
4.  Run any of the batch processes (e.g., `RunResumableBatch`) from the Excel macro menu. The framework will automatically bind to the active terminal session (e.g., `pw3270:A`).

## 🤝 Contributing
Contributions, issues, and feature requests are welcome! 
If you have suggestions for improving the codebase, please:
1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## ⚠️ Disclaimer
This toolkit interacts with live enterprise systems and processes sensitive financial and personal data (PII). It is intended strictly for authorized use by system operators. The creators of this repository assume no liability for misuse, unintended data modification, or system lockouts resulting from the execution of these scripts.

## 📄 License
Distributed under the MIT License. 

```text
MIT License

Copyright (c) 2026 [Your Name/GitHub Handle]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.