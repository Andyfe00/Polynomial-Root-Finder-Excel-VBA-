# Polynomial-Root-Finder-Excel-VBA-
This project implements an interactive numerical analysis tool in **Excel with VBA macros** to compute the **zeros (roots) of polynomial functions** using two classical numerical methods:

- **Bisection Method**
- **Newton–Raphson Method**

It was designed as an educational tool to visualize convergence and iterative behavior of these algorithms directly within Excel, with dynamic interface creation, error tracking, and step-by-step tables.

---

## ⚙️ Overview

The file **`Ceros De Funciones.xlsm`** contains a complete VBA-based system that:
- Builds input tables dynamically based on the polynomial degree.
- Allows the user to select the numerical method.
- Computes all iterations automatically until the error tolerance is met.
- Displays detailed tables and iteration results directly in the worksheet.

---

## 📂 Files

| File | Description |
|------|--------------|
| `Root Finder.xlsm` | Excel workbook with macros for polynomial root finding. |
| `README.md` | Documentation and usage guide. |

---

## 🧰 Requirements

- Microsoft **Excel (Windows)** with **macros enabled** (`.xlsm` support).
- VBA macros allowed (you must **enable content** when opening the file).

---

## 🚀 How to Use

1. **Open the workbook** `Root Finder.xlsm`.
2. If prompted, click **“Enable Content”** to activate macros.
3. In cell **D5**, enter the **degree of the polynomial** (e.g., `3` for a cubic polynomial).
4. Click the **“Leer”** button — the macro will:
   - Create a table for coefficient input (columns labeled `x^0`, `x^1`, …, `x^n`).
   - Automatically add buttons for each method: **Bisección**, **Newton**, and **Leer** (refresh).
5. Enter the polynomial coefficients in the new table.
6. Choose your method:
   - **Bisection** → specify interval `[a, b]` and error tolerance.
   - **Newton** → specify initial point and error tolerance.
7. Click **“Calcule”** to run the method.
8. The program will automatically generate an **iteration table** showing all steps, function values, and errors.

> ⚠️ Academic Disclaimer  
> This project was developed by **Andrés Felipe Porras** as part of the  
> Metodos Numericos module at **Universidad Militar Nueva Granada (2017–2021)**.  
> The content is shared for educational and portfolio purposes only.  
> No proprietary or restricted university materials are included.

