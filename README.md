# LibreOffice PPT-Style Math Workflow 🚀

[![License: MIT](https://shields.io)](https://opensource.org)
[![Platform: Linux](https://shields.io)](https://linuxmint.com)
[![Powered by Espanso](https://shields.io)](https://espanso.org)

Bring Microsoft PowerPoint's fast, "type-in-the-flow" equation experience (`Alt + =` and Math AutoCorrect) to **LibreOffice Impress** and **LibreOffice Writer** on Linux.

> **Why this exists:** LibreOffice Impress relies on separate OLE Formula objects that interrupt typing flow. This project combines a collision-proof **Espanso** Math AutoCorrect dictionary for instant inline Unicode math (`\alpha^2 + \beta^2 → \infty`) with a **TexMaths** workflow for crisp, re-editable 2D stacked equations (`\dfrac{dy}{dx}`, integrals, and matrices).

---

## ✨ Features

* **Instant Inline Math AutoCorrect:** Type `\alpha`, `\beta`, `\sum`, `\neq`, `->`, `^2`, or `_{i+1}` directly inside any Impress text box or Writer paragraph.
* **Collision-Proof Triggers:** Uses word boundaries on backslash commands so typing `\in` never prematurely triggers before `\infty` or `\int`.
* **Multi-Character Super/Subscripts:** Supports regex-powered expressions like `e^{2n+1}` → `e²ⁿ⁺¹` and `x_{i+1}` → `xᵢ₊₁`.
* **True 2D Stacked Equations:** Pairs seamlessly with **TexMaths** (`Ctrl + =`) for high-resolution, re-editable SVG fractions, matrices, and calculus notation.
* **Zero LaTeX Hijacking:** Configured so inline triggers don't interfere when typing raw LaTeX inside TexMaths.

---

## 📸 Demo

*(Tip: Record a short screen capture using `Ctrl + Alt + Shift + R` on Linux Mint, convert to GIF, and replace this line with `![Demo](demo.gif)`)*

---

## 🛠️ Installation

### 1. Prerequisites
Ensure you have **Espanso** installed on your Linux system:
* Follow the [official Espanso Linux installation guide](https://espanso.orgdocs/install/linux/) (choose the X11 `.deb` package for Linux Mint / Ubuntu X11).

### 2. Install the Math AutoCorrect Rules
Clone this repository and copy `ppt_math.yml` into your Espanso match directory:

```bash
git clone https://github.com
cd libreoffice-ppt-math
mkdir -p ~/.config/espanso/match
cp config/ppt_math.yml ~/.config/espanso/match/
espanso restart
```

*(Optional)* If you want to customize or add new symbols, edit `scripts/generate_rules.py` and run:
```bash
python3 scripts/generate_rules.py
espanso restart
```

---

## 📐 Setting Up Stacked 2D Equations (TexMaths)

For tall fractions (`\dfrac{dy}{dx}`), integrals, and matrices that cannot be represented on a single text line:

1. **Install the lightweight LaTeX backend (Ubuntu / Linux Mint / Debian):**
   ```bash
   sudo apt update && sudo apt install texlive-latex-base dvipng dvisvgm
   ```
2. **Install the TexMaths Extension:**
   * Download `TexMaths-0.52.oxt` (or newer) from the [TexMaths SourceForge page](https://sourceforge.net).
   * Open LibreOffice, go to **Tools > Extension Manager > Add**, select the `.oxt` file, and restart LibreOffice.
3. **Bind a Keyboard Shortcut (`Ctrl + =`):**
   * Go to **Tools > Customize > Keyboard**.
   * Select **`Ctrl + =`** in the top shortcut list (and choose **LibreOffice** in the top-right radio button to apply globally across Impress and Writer).
   * Under **Category**, expand **Application Macros > user > TexMaths > TexMathsEquations** and select **`main`**.
   * Click **Modify**, then **OK**.
4. **Recommended TexMaths Settings:**
   * Click the **π** icon on your toolbar, open **System Configuration**, and set the default image format to **SVG** for infinitely sharp scaling.

---

## ⌨️ Quick Cheat Sheet

### Inline Typing (Works Anywhere)

| You Type | Result | Notes |
| :--- | :---: | :--- |
| `\alpha ` / `\Omega ` | `α` / `Ω` | Press `Space` after backslash commands |
| `\infty ` / `\int ` | `∞` / `∫` | Collision-safe (`\in ` won't hijack it) |
| `^2` / `^{2n+1}` | `²` / `²ⁿ⁺¹` | Instant single or multi-char superscripts |
| `_0` / `_{i+1}` | `₀` / `ᵢ₊₁` | Instant single or multi-char subscripts |
| `->` / `=>` / `!=` | `→` / `⇒` / `≠` | Instant operators and arrows |

### 2D Stacked Equations (TexMaths)
* **In Impress:** Press `Esc` twice (to exit any text box), press `Ctrl + =`, type `\dfrac{dy}{dx}`, and press `Enter`.
* **In Writer:** Press `Ctrl + =` directly inside a paragraph for true inline baseline-aligned equations.
* **To Edit Later:** **Single-click** any equation on your slide/page and press `Ctrl + =` to reopen and edit its LaTeX source.

---

## 🤝 Contributing

Contributions are welcome! If you'd like to add symbol sets for **Chemistry** (`\rightleftharpoons`), **Physics** (`\hbar`, `\vec{F}`), or **Logic/Set Theory**:
1. Fork the repository.
2. Add your triggers to `scripts/generate_rules.py` and regenerate `config/ppt_math.yml`.
3. Open a Pull Request describing the new shortcuts.

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.
