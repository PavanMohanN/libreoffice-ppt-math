# libreoffice-ppt-math

### Demo

https://github.com/user-attachments/assets/a69e2a02-92cf-45cd-a3bb-c1a4e5b2e0c1

*Typing `\alpha^2 + \beta^2 -> \infty` natively in LibreOffice Impress. *AI Generated *




Type math in **LibreOffice Impress & Writer** just like PowerPoint's Math AutoCorrect (`\alpha` → `α`, `^2` → `²`, `^{n+1}` → `ⁿ⁺¹`, `->` → `→`).

---

## Step 1: Install Required Tools (Linux)

You need three free tools before running this project:

1. **Espanso** *(the text-expander that converts shortcuts into math symbols)*
   * Follow the official 2-minute guide: **[Install Espanso on Linux](https://espanso.org/docs/install/linux/)**
   * Make sure it is running by typing `espanso status` in your terminal.
2. **Git** *(to download this project)*
   * Ubuntu/Debian/Mint: `sudo apt install git` (or see **[Git Downloads](https://git-scm.com/downloads)**)
3. **Python 3** *(for multi-character superscripts/subscripts)*
   * Usually pre-installed on Linux. Check with `python3 --version` (or see **[Python Downloads](https://www.python.org/downloads/)**)

---

## Step 2: Install This Project

Open your terminal and copy-paste these three lines:

```bash
git clone [https://github.com/PavanMohanN/libreoffice-ppt-math.git](https://github.com/PavanMohanN/libreoffice-ppt-math.git)
cd libreoffice-ppt-math
bash install.sh
