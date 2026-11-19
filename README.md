# libreoffice-ppt-math

### Demo

https://github.com/user-attachments/assets/a69e2a02-92cf-45cd-a3bb-c1a4e5b2e0c1

*Typing `\alpha^2 + \beta^2 -> \infty` natively in LibreOffice Impress. *AI Generated *


Type math natively in **LibreOffice Impress & Writer** just like PowerPoint's Math AutoCorrect (`\alpha` → `α`, `^2` → `²`, `^{n+1}` → `ⁿ⁺¹`, `->` → `→`), powered by [Espanso](https://espanso.org).

## Demo

https://github.com/user-attachments/assets/04532d7d-3800-45fc-a8ed-c40ffbd63c70

*Typing `\alpha^2 + \beta^2 -> \infty` natively in LibreOffice Impress.*

---

## Step 1: Install Prerequisites

Before installing, make sure you have these three free tools:

1. **Espanso** *(the text-expander that converts shortcuts into symbols)*
   * Install via the official guide: **[Espanso Installation Docs](https://espanso.org/docs/install/)** (Linux, macOS, Windows)
   * Verify it is running: `espanso status`
2. **Git** *(to download this repository)*
   * Ubuntu/Debian/Mint: `sudo apt install git` (or **[Git Downloads](https://git-scm.com/downloads)**)
3. **Python 3** *(for multi-character superscripts/subscripts and rule generation)*
   * Usually pre-installed on Linux/macOS. Check with `python3 --version` (or **[Python Downloads](https://www.python.org/downloads/)**)

---

## Step 2: Install

Open your terminal and run:

```bash
git clone [https://github.com/PavanMohanN/libreoffice-ppt-math.git](https://github.com/PavanMohanN/libreoffice-ppt-math.git)
cd libreoffice-ppt-math

```

Choose **one** of the two installation modes:

* **Option A — Global Install** *(works in LibreOffice and all other apps)*:
```bash
bash install.sh

```


* **Option B — LibreOffice Only** *(recommended for developers so `->`, `!=`, and `1/2` don't trigger in code editors or terminals)*:
```bash
bash install.sh --libreoffice-only

```



*(The installer automatically detects your Espanso folder, backs up any existing config, installs the rules, and restarts Espanso.)*

---

## Step 3: Try It Out

Open **LibreOffice Impress** (or Writer) and type:

```text
\alpha^2 + \beta^2 -> \infty 

```

*(Press **Space** after backslash commands like `\alpha` and `\infty`. Arrows, operators, accents, and super/subscripts convert instantly.)*

---

## Cheat Sheet

| Category | What You Type | Result | Trigger |
| --- | --- | --- | --- |
| **Greek Letters** | `\alpha`, `\beta`, `\theta`, `\Omega` | `α`, `β`, `θ`, `Ω` | `+ Space` |
| **Calculus & Symbols** | `\sum`, `\int`, `\partial`, `\infty`, `\pm`, `\degree` | `∑`, `∫`, `∂`, `∞`, `±`, `°` | `+ Space` |
| **Sets & Logic** | `\R`, `\C`, `\N`, `\Z`, `\in`, `\forall`, `\exists`, `\therefore` | `ℝ`, `ℂ`, `ℕ`, `ℤ`, `∈`, `∀`, `∃`, `∴` | `+ Space` |
| **Arrows & Operators** | `->`, `=>`, `<->`, `!=`, `<=`, `>=`, `~=` | `→`, `⇒`, `↔`, `≠`, `≤`, `≥`, `≈` | Instant |
| **Common Fractions** | `1/2`, `1/3`, `2/3`, `1/4`, `3/4` | `½`, `⅓`, `⅔`, `¼`, `¾` | Instant |
| **Single Super/Subscripts** | `^2`, `^n`, `^-`, `_0`, `_i`, `_x` | `²`, `ⁿ`, `⁻`, `₀`, `ᵢ`, `ₓ` | Instant |
| **Multi Super/Subscripts** | `^{2n+1}`, `_{i+1}`, `^{(k)}` | `²ⁿ⁺¹`, `ᵢ₊₁`, `⁽ᵏ⁾` | Instant on `}` |
| **Accents** | `\vec{v}`, `\bar{x}`, `\hat{x}`, `\tilde{x}` | `v⃗`, `x̄`, `x̂`, `x̃` | Instant on `}` |

---

## Customizing Rules

Edit `scripts/generate_rules.py` to add or modify shortcuts, then rebuild and reinstall:

```bash
python3 scripts/generate_rules.py
bash install.sh

```

---

## Uninstall

To cleanly remove all installed math rules (your `.bak.*` backups are left untouched):

```bash
bash uninstall.sh

```

---

## Known Limitations & Tips

* **Unicode Subscript Gaps:** The Unicode standard does not include subscript characters for `b, c, d, f, g, q, w, y, z` or uppercase letters. Multi-character subscripts (`_{...}`) only expand supported characters (`0-9`, `+ - = ( )`, `a e h i j k l m n o p r s t u v x`).
* **Best Fonts for Accents:** Combining accents (`\vec{v}`, `\bar{x}`, `\hat{x}`, `\tilde{x}`) align cleanest in **Liberation Sans**, **DejaVu Sans**, or **Noto Sans**.

---

## License

MIT License — see [LICENSE](https://www.google.com/search?q=LICENSE) for details.

```
**One quick thing before you push:** Run `python3 scripts/generate_rules.py` once so your repo's `config/ppt_math.yml` includes the newly added `\R`, `\C`, `\N`, `\Z`, `\degree`, `\hat`, and `\tilde` rules—otherwise Option A (`bash install.sh`) will copy the older file.

```
