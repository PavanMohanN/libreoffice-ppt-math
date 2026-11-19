# LibreOffice PPT-Style Math (`libreoffice-ppt-math`)

Type mathematical equations natively in **LibreOffice Impress** (and Writer) using PowerPoint-style Math AutoCorrect shortcuts (e.g., `\alpha` → `α`, `^2` → `²`, `^{n+1}` → `ⁿ⁺¹`, `->` → `→`), powered by [Espanso](https://espanso.org)[cite: 1, 3].

---

## Prerequisites

1. **Espanso** installed and running (`espanso --version`).
2. **Python 3** (`python3 --version`) — required for multi-character superscript/subscript regex expansions and regenerating rules[cite: 1, 2].

---

## Quick Start (Reproduce in 3 Steps)

### 1. Clone the repository
```bash
git clone git@github.com:PavanMohanN/libreoffice-ppt-math.git
cd libreoffice-ppt-math

```

### 2. Run the installer

```bash
bash install.sh

```

The installer will:

* Verify that `espanso` is installed.


* Back up any existing `~/.config/espanso/match/ppt_math.yml` with a timestamp.


* Copy `config/ppt_math.yml` to `~/.config/espanso/match/ppt_math.yml`.


* Restart Espanso to load the new rules.



> **Manual / Direct Python Install:** You can also generate and install the config directly via:
> ```bash
> python3 scripts/generate_rules.py --install
> espanso restart
> 
> ```
> 
> 
> 
> 
> 

### 3. Test it

Open LibreOffice Impress (or any text field) and type:

```text
\alpha^2 + \beta^2 -> \infty 

```


(Note: Backslash commands require a trailing space or word boundary to expand; operators and single-character super/subscripts expand immediately.)

---

## Supported Syntax

| Category | Trigger Examples | Output | Behavior |
| --- | --- | --- | --- |
| **Greek Letters** | `\alpha`, `\theta`, `\Omega`

 | `α`, `θ`, `Ω`

 | Word-boundary (`Space`)

 |
| **Operators & Calculus** | `\times`, `\pm`, `\partial`, `\sum`, `\int`, `\infty`

 | `×`, `±`, `∂`, `∑`, `∫`, `∞`

 | Word-boundary (`Space`)

 |
| **Logic & Sets** | `\forall`, `\exists`, `\in`, `\subset`, `\cup`, `\therefore`

 | `∀`, `∃`, `∈`, `⊂`, `∪`, `∴`

 | Word-boundary (`Space`)

 |
| **Instant Operators & Arrows** | `!=`, `<=`, `>=`, `~=`, `->`, `<=>`, `=>`

 | `≠`, `≤`, `≥`, `≈`, `→`, `⇔`, `⇒`

 | Instant

 |
| **Common Fractions** | `1/2`, `1/3`, `2/3`, `1/4`, `3/4`

 | `½`, `⅓`, `⅔`, `¼`, `¾`

 | Instant

 |
| **Single Super/Subscripts** | `^2`, `^n`, `^-`, `_0`, `_i`, `_x`

 | `²`, `ⁿ`, `⁻`, `₀`, `ᵢ`, `ₓ`

 | Instant

 |
| **Multi-Char Super/Subscripts** | `^{2n+1}`, `_{i+1}`, `^{(k)}` | `²ⁿ⁺¹`, `ᵢ₊₁`, `⁽ᵏ⁾` | Instant on closing `}` |
| **Vector & Overbar Accents** | `\vec{v}`, `\bar{x}` | `v⃗`, `x̄` | Instant on closing `}` |

---

## Customizing Rules

To add or modify triggers, edit `scripts/generate_rules.py` and rebuild `config/ppt_math.yml`:

```bash
python3 scripts/generate_rules.py
bash install.sh

```



---

## Known Limitations & Tips

* **Unicode Subscript Coverage:** The Unicode standard lacks subscript glyphs for `b, c, d, f, g, q, w, y, z` and uppercase letters. Multi-character subscripts (`_{...}`) only match supported characters (`0-9`, `+ - = ( )`, `a e h i j k l m n o p r s t u v x`).


* **Font Rendering for Accents:** Combining diacritics (`\vec{v}` and `\bar{x}`) align best in fonts with strong Unicode math coverage such as **Liberation Sans**, **DejaVu Sans**, or **Noto Sans**.


* **Scoping to LibreOffice Only (Optional):** To prevent instant triggers (`->`, `1/2`, `<=`) from firing in code editors or terminals, wrap your installed `ppt_math.yml` in an [Espanso app-specific configuration](https://espanso.org/docs/configuration/app-specific-configurations/) filtering for `soffice.bin`.


* **macOS / Windows Paths:** `install.sh` targets Linux (`~/.config/espanso/match`). On macOS or Windows, copy `config/ppt_math.yml` into the `match/` folder shown by `espanso path`.



```

Do you want to share your existing `README.md` if you'd like me to merge sections, or are you ready to commit this?

```
