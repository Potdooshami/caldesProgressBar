#  CALDES Slide Indicator (VBA)

This repository provides a VBA macro tool that creates a custom slide indicator for PowerPoint presentations, featuring the **CALDES (Center for Artificial Low Dimensional Electronic Systems)** identity.

Instead of standard page numbers, the CALDES logo (Soliton) moves across the slide to visually represent the progress of the presentation.

## 1. The Meaning of the Logo: Soliton
<img width="266" height="83" alt="image" src="https://github.com/user-attachments/assets/5b7b5bea-2372-4f10-a3d7-1a0e8849a56f" />




The CALDES logo represents a **Soliton**.

* **Definition:** A soliton is a self-reinforcing solitary wave packet that maintains its shape while propagating at a constant velocity. It is a specific solution to non-linear wave equations.
* **Symbolism:** In the context of our research, it represents unique **topological excitations** in low-dimensional electronic systems, symbolizing the robustness and structural stability of the physical phenomena we study.

Just as a soliton propagates through a medium without energy dissipation, this indicator animates the logo traveling from left to right as the presentation proceeds.

## 2. Features

[![ad86rj](https://github.com/user-attachments/assets/8170daaa-d43b-4617-a395-115cd219ea85)](https://caldes.ibs.re.kr/html/caldes_en/)

* **Dynamic Positioning:** Automatically calculates the logo's position based on the total slide count and current slide number.
* **Visual Progress Bar:** Offers an intuitive visual cue for the audience regarding the remaining time/pages.
* **Customizable:** The VBA script allows easy adjustment of the logo's size, vertical position, and transparency.

## 3. How to Apply the VBA Code

This tool requires enabling Macros in PowerPoint. Follow these steps:

### Step 1: Enable the Developer Tab
1.  Open PowerPoint.
2.  Go to **File** > **Options** > **Customize Ribbon**.
3.  In the right-hand list, check the box for **[Developer]** and click **OK**.

### Step 2: Insert the VBA Code
1.  Press `Alt` + `F11` to open the VBA Editor.
2.  Go to **Insert** > **Module** in the top menu.
3.  Copy the code from `script.vba` in this repository and paste it into the module window.

### Step 3: Run the Macro
1.  Press `Alt` + `F8` in PowerPoint.
2.  Select the macro (e.g., `AddSolitonIndicator`) and click **[Run]**.
3.  Verify that the logo indicator appears correctly at the bottom of your slides.

### Step 4: Save
* Save the file as a **PowerPoint Macro-Enabled Presentation (`.pptm`)** to ensure the code remains active.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
