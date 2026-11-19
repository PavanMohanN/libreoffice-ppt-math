# LibreOffice PPT-Style Math Workflow 🚀

Bring Microsoft PowerPoint-style math shortcuts and AutoCorrect behavior to **LibreOffice Impress** and **Writer** on Linux using Espanso and TexMaths [PavanMohanN/libreoffice-ppt-math].

---

## ✨ What You Get
* **Mid-Sentence Math:** Type standard shortcuts like `\alpha`, `\beta`, `\sum`, or `^2` directly into documents with automatic conversion [PavanMohanN/libreoffice-ppt-math].
* **Stacked 2D Equations:** Combine with TexMaths and `Ctrl + =` to render re-editable LaTeX expressions [PavanMohanN/libreoffice-ppt-math].

---

## 🛠️ Step-by-Step Setup Guide

### Step 1: Install Espanso
Register and start the Espanso service on X11/Linux Mint/Ubuntu systems (refer to the [official Espanso docs](https://espanso.org)) [PavanMohanN/libreoffice-ppt-math]:
```bash
wget https://github.com
sudo apt install ./espanso-debian-x11-amd64.deb
rm espanso-debian-x11-amd64.deb
espanso service register
espanso start
```

### Step 2: Install the Math Shortcut Pack
Clone the repository and run the installation script [PavanMohanN/libreoffice-ppt-math]:
```bash
git clone https://github.com
cd libreoffice-ppt-math
bash install.sh
```

### Step 3: Set Up Stacked Fractions & TexMaths (Optional)
1. Install LaTeX dependencies and the **TexMaths** extension via LibreOffice **Tools > Extension Manager** [PavanMohanN/libreoffice-ppt-math].
2. Bind `Ctrl + =` to `TexMathsEquations -> main` and switch the output format to **SVG** [PavanMohanN/libreoffice-ppt-math].

---

## ⌨️ Quick Cheat Sheet
* **Inline Examples:** `\alpha ` $\rightarrow$ `α`, `->` $\rightarrow$ `→`, `x^2` $\rightarrow$ `x²` [PavanMohanN/libreoffice-ppt-math].
* **Stacked Equations:** Press `Ctrl + =`, type your LaTeX expression, and press `Enter` [PavanMohanN/libreoffice-ppt-math].

For the complete set of scripts and setup instructions, visit the [GitHub repository](https://github.com).

---

## 📄 License
MIT License — free to use, modify, and share [PavanMohanN/libreoffice-ppt-math].
