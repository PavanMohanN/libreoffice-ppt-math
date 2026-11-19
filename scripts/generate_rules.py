#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
PowerPoint-style Math AutoCorrect for Espanso -- generator & installer
======================================================================

HOW YOU TYPE (one rule, no exceptions)
    Type the expression in LaTeX / PowerPoint "linear format" WITHOUT spaces,
    then press SPACE. The whole expression is converted in one go, exactly like
    PowerPoint builds up an equation when you press space.

        \alpha             ->  α              x^2              ->  x²
        \int_0^1           ->  ∫₀¹            \sum_{i=1}^{n}x_i ->  ∑ᵢ₌₁ⁿ xᵢ
        \lim_{x\to0}       ->  lim_(x→0)      \sqrt{x^2+1}     ->  √x̅²̅+̅1̅
        \frac{1}{3}        ->  ⅓              \frac{a+b}{c}    ->  (a+b)/c
        A\cup{B}           ->  A∪B            \overline{AB}    ->  A̅B̅
        a<=b  a!=b  a->b   ->  a≤b  a≠b  a→b  <=>  <->  ==>    ->  ⇔  ↔  ⟹

    * Need a space inside an expression?  Use ~  (LaTeX tie) or \, \; \quad
    * Backspace right after a conversion restores what you typed (Espanso undo).
    * Characters with no Unicode super/subscript form (e.g. subscript y, ∞, →)
      fall back to readable linear format:  v_y,  ∫_(−∞)^∞,  lim_(x→0).

USAGE
    python3 generate_rules.py                          # write package.yml (repo)
    python3 generate_rules.py --install                # install for every app
    python3 generate_rules.py --install --libreoffice-only
    python3 generate_rules.py --uninstall              # remove installed files
    python3 generate_rules.py --check                  # diagnose conflicts
    python3 generate_rules.py --self-test              # run the built-in tests
    python3 generate_rules.py --try "\int_0^1" "x^{-1}"  # preview conversions
    python3 generate_rules.py --list [text]            # list every command

ADDING YOUR OWN SYMBOLS
    Edit the USER_SYMBOLS / USER_FUNCTIONS / USER_OPERATORS blocks just below,
    then re-run this script with the same flags you used to install.
"""

import argparse
import base64
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import zipfile
import zlib
from pathlib import Path

# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  USER-EDITABLE SECTION -- add your own symbols here, then re-run.        ║
# ╚══════════════════════════════════════════════════════════════════════════╝

# \command  ->  replacement text.  Names are letters only (A-Z, a-z).
# A leading backslash is optional. Entries here override the built-in tables.
USER_SYMBOLS = {
    # "\\bigO": "𝒪",
    # "\\eps": "ε",
    # "\\ohm": "Ω",
}

# Commands rendered as upright operator names, like \sin -> sin.
USER_FUNCTIONS = [
    # "csch",
]

# Plain-ASCII sequences converted inside a math expression (longest match wins).
# Must be 2-4 printable ASCII characters, no letters, digits or spaces.
USER_OPERATORS = {
    # "|=": "⊨",
}

SETTINGS = {
    # "vinculum": √x̅+̅1̅ (bar over the whole radicand)  |  "parens": √(x+1)
    "sqrt_style": "vinculum",
    # Turn a typed hyphen into a true minus sign (−) inside math expressions.
    "minus_sign": True,
    # Thin space after sin/log/lim/∑/∫ when a letter or digit follows (LaTeX style).
    "thin_space_after_operators": True,
}

# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  BUILT-IN TABLES                                                         ║
# ╚══════════════════════════════════════════════════════════════════════════╝

GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ϵ",
    "varepsilon": "ε", "zeta": "ζ", "eta": "η", "theta": "θ", "vartheta": "ϑ",
    "iota": "ι", "kappa": "κ", "varkappa": "ϰ", "lambda": "λ", "mu": "μ",
    "nu": "ν", "xi": "ξ", "omicron": "ο", "pi": "π", "varpi": "ϖ", "rho": "ρ",
    "varrho": "ϱ", "sigma": "σ", "varsigma": "ς", "tau": "τ", "upsilon": "υ",
    "phi": "ϕ", "varphi": "φ", "chi": "χ", "psi": "ψ", "omega": "ω",
    "digamma": "ϝ",
    "Alpha": "Α", "Beta": "Β", "Gamma": "Γ", "Delta": "Δ", "Epsilon": "Ε",
    "Zeta": "Ζ", "Eta": "Η", "Theta": "Θ", "varTheta": "ϴ", "Iota": "Ι",
    "Kappa": "Κ", "Lambda": "Λ", "Mu": "Μ", "Nu": "Ν", "Xi": "Ξ",
    "Omicron": "Ο", "Pi": "Π", "Rho": "Ρ", "Sigma": "Σ", "Tau": "Τ",
    "Upsilon": "Υ", "Phi": "Φ", "Chi": "Χ", "Psi": "Ψ", "Omega": "Ω",
    "Digamma": "Ϝ",
}

HEBREW = {
    "aleph": "ℵ", "beth": "ℶ", "bet": "ℶ", "gimel": "ℷ", "daleth": "ℸ", "dalet": "ℸ",
}

BINARY_OPERATORS = {
    "pm": "±", "mp": "∓", "times": "×", "div": "÷", "cdot": "⋅", "centerdot": "·",
    "bullet": "∙", "ast": "∗", "star": "⋆", "circ": "∘", "bigcirc": "◯",
    "oplus": "⊕", "ominus": "⊖", "otimes": "⊗", "oslash": "⊘", "odot": "⊙",
    "circledast": "⊛", "circledcirc": "⊚", "circleddash": "⊝",
    "boxplus": "⊞", "boxminus": "⊟", "boxtimes": "⊠", "boxdot": "⊡",
    "dagger": "†", "ddagger": "‡", "dag": "†", "ddag": "‡", "amalg": "⨿",
    "wr": "≀", "setminus": "∖", "smallsetminus": "∖", "dotplus": "∔",
    "ltimes": "⋉", "rtimes": "⋊", "divideontimes": "⋇", "intercal": "⊺",
    "minus": "−", "cap": "∩", "cup": "∪", "Cap": "⋒", "Cup": "⋓",
    "sqcap": "⊓", "sqcup": "⊔", "uplus": "⊎", "wedge": "∧", "vee": "∨",
    "curlywedge": "⋏", "curlyvee": "⋎", "barwedge": "⊼", "veebar": "⊻",
    "doublebarwedge": "⩞", "triangleleft": "◁", "triangleright": "▷",
    "bigtriangleup": "△", "bigtriangledown": "▽", "lhd": "⊲", "rhd": "⊳",
    "unlhd": "⊴", "unrhd": "⊵", "ldiv": "∕", "sdiv": "⁄",
    "xor": "⊕", "nand": "⊼", "nor": "⊽",
}

RELATIONS = {
    "ne": "≠", "neq": "≠", "le": "≤", "leq": "≤", "ge": "≥", "geq": "≥",
    "leqq": "≦", "geqq": "≧", "leqslant": "⩽", "geqslant": "⩾",
    "ll": "≪", "gg": "≫", "lll": "⋘", "ggg": "⋙",
    "nless": "≮", "ngtr": "≯", "nleq": "≰", "ngeq": "≱",
    "lneq": "⪇", "gneq": "⪈", "lneqq": "≨", "gneqq": "≩",
    "lesssim": "≲", "gtrsim": "≳", "lessgtr": "≶", "gtrless": "≷",
    "lessdot": "⋖", "gtrdot": "⋗",
    "approx": "≈", "approxeq": "≊", "napprox": "≉", "sim": "∼", "nsim": "≁",
    "simeq": "≃", "nsimeq": "≄", "backsim": "∽", "eqsim": "≂",
    "cong": "≅", "ncong": "≇", "equiv": "≡", "nequiv": "≢",
    "propto": "∝", "varpropto": "∝",
    "doteq": "≐", "doteqdot": "≑", "Doteq": "≑", "fallingdotseq": "≒",
    "risingdotseq": "≓", "coloneqq": "≔", "coloneq": "≔", "eqqcolon": "≕",
    "eqcolon": "≕", "defeq": "≝", "triangleq": "≜", "Deltaeq": "≜",
    "measeq": "≞", "questeq": "≟", "stareq": "≛", "wedgeq": "≙", "veeeq": "≚",
    "eqcirc": "≖", "circeq": "≗", "bumpeq": "≏", "Bumpeq": "≎", "asymp": "≍",
    "between": "≬",
    "prec": "≺", "succ": "≻", "preceq": "⪯", "succeq": "⪰",
    "preccurlyeq": "≼", "succcurlyeq": "≽", "nprec": "⊀", "nsucc": "⊁",
    "precsim": "≾", "succsim": "≿",
    "perp": "⊥", "parallel": "∥", "nparallel": "∦", "mid": "∣", "nmid": "∤",
    "divides": "∣", "smile": "⌣", "frown": "⌢",
    "vdash": "⊢", "dashv": "⊣", "models": "⊨", "vDash": "⊨", "Vdash": "⊩",
    "VDash": "⊫", "Vvdash": "⊪", "nvdash": "⊬", "nvDash": "⊭", "nVdash": "⊮",
    "nVDash": "⊯", "pitchfork": "⋔", "multimap": "⊸", "bowtie": "⋈",
    "Join": "⨝", "ratio": "∶", "Colon": "∷", "dblcolon": "∷",
    "therefore": "∴", "because": "∵", "indep": "⫫",
}

SETS = {
    "in": "∈", "notin": "∉", "ni": "∋", "owns": "∋", "notni": "∌", "nni": "∌",
    "notcontain": "∌", "notelement": "∉",
    "subset": "⊂", "supset": "⊃", "subseteq": "⊆", "supseteq": "⊇",
    "subsetneq": "⊊", "supsetneq": "⊋", "nsubset": "⊄", "nsupset": "⊅",
    "nsubseteq": "⊈", "nsupseteq": "⊉", "subseteqq": "⫅", "supseteqq": "⫆",
    "Subset": "⋐", "Supset": "⋑", "sqsubset": "⊏", "sqsupset": "⊐",
    "sqsubseteq": "⊑", "sqsupseteq": "⊒", "superset": "⊃", "superseteq": "⊇",
    "sqsuperseteq": "⊒", "emptyset": "∅", "varnothing": "∅",
    "complement": "∁", "powerset": "𝒫", "symdiff": "∆", "union": "∪",
    "intersection": "∩",
    "N": "ℕ", "Z": "ℤ", "Q": "ℚ", "R": "ℝ", "C": "ℂ",
    "naturals": "ℕ", "integers": "ℤ", "rationals": "ℚ", "reals": "ℝ",
    "complexes": "ℂ", "Bbbk": "𝕜",
}

LOGIC = {
    "forall": "∀", "exists": "∃", "nexists": "∄", "neg": "¬", "lnot": "¬",
    "land": "∧", "lor": "∨", "top": "⊤", "bot": "⊥", "Box": "□",
    "Diamond": "◇", "qed": "∎", "blacksquare": "■", "checkmark": "✓",
}

ARROWS = {
    "to": "→", "gets": "←", "rightarrow": "→", "leftarrow": "←",
    "uparrow": "↑", "downarrow": "↓", "leftrightarrow": "↔", "updownarrow": "↕",
    "Rightarrow": "⇒", "Leftarrow": "⇐", "Uparrow": "⇑", "Downarrow": "⇓",
    "Leftrightarrow": "⇔", "Updownarrow": "⇕",
    "longrightarrow": "⟶", "longleftarrow": "⟵", "longleftrightarrow": "⟷",
    "Longrightarrow": "⟹", "Longleftarrow": "⟸", "Longleftrightarrow": "⟺",
    "implies": "⇒", "impliedby": "⇐", "iff": "⇔",
    "mapsto": "↦", "longmapsto": "⟼", "mapsfrom": "↤", "longmapsfrom": "⟻",
    "hookrightarrow": "↪", "hookleftarrow": "↩",
    "nearrow": "↗", "searrow": "↘", "swarrow": "↙", "nwarrow": "↖",
    "rightharpoonup": "⇀", "rightharpoondown": "⇁", "leftharpoonup": "↼",
    "leftharpoondown": "↽", "upharpoonleft": "↿", "upharpoonright": "↾",
    "downharpoonleft": "⇃", "downharpoonright": "⇂",
    "rightleftharpoons": "⇌", "leftrightharpoons": "⇋", "lrhar": "⇋", "rlhar": "⇌",
    "rightrightarrows": "⇉", "leftleftarrows": "⇇", "rightleftarrows": "⇄",
    "leftrightarrows": "⇆", "upuparrows": "⇈", "downdownarrows": "⇊",
    "twoheadrightarrow": "↠", "twoheadleftarrow": "↞", "rightarrowtail": "↣",
    "leftarrowtail": "↢", "looparrowright": "↬", "looparrowleft": "↫",
    "curvearrowright": "↷", "curvearrowleft": "↶", "circlearrowright": "↻",
    "circlearrowleft": "↺", "Lsh": "↰", "Rsh": "↱", "rightsquigarrow": "⇝",
    "leadsto": "⇝", "leftrightsquigarrow": "↭",
    "nrightarrow": "↛", "nleftarrow": "↚", "nleftrightarrow": "↮",
    "nRightarrow": "⇏", "nLeftarrow": "⇍", "nLeftrightarrow": "⇎",
    "Rrightarrow": "⇛", "Lleftarrow": "⇚", "dashrightarrow": "⇢",
    "dashleftarrow": "⇠",
}

# Symbols that take limits (\sum_{i=1}^{n}, \int_a^b, \bigcup_i ...).
BIG_OPERATORS = {
    "sum": "∑", "prod": "∏", "coprod": "∐", "int": "∫", "iint": "∬",
    "iiint": "∭", "iiiint": "⨌", "oint": "∮", "oiint": "∯", "oiiint": "∰",
    "intclockwise": "∱", "varointclockwise": "∲", "ointctrclockwise": "∳",
    "coint": "∲", "aoint": "∳", "fint": "⨏",
    "bigcup": "⋃", "bigcap": "⋂", "bigvee": "⋁", "bigwedge": "⋀",
    "bigoplus": "⨁", "bigotimes": "⨂", "bigodot": "⨀", "biguplus": "⨄",
    "bigsqcup": "⨆", "bigsqcap": "⨅",
}

CALCULUS_AND_LETTERLIKE = {
    "infty": "∞", "infinity": "∞", "partial": "∂", "nabla": "∇",
    "increment": "∆", "inc": "∆",
    "dd": "ⅆ", "Dd": "ⅅ", "ee": "ⅇ", "ii": "ⅈ", "jj": "ⅉ",
    "prime": "′", "dprime": "″", "trprime": "‴", "qprime": "⁗",
    "pprime": "″", "ppprime": "‴", "pppprime": "⁗", "backprime": "‵",
    "hbar": "ℏ", "hslash": "ℏ", "ell": "ℓ", "wp": "℘", "Re": "ℜ", "Im": "ℑ",
    "imath": "ı", "jmath": "ȷ", "eth": "ð", "mho": "℧", "Finv": "Ⅎ", "Game": "⅁",
    "surd": "√", "degree": "°", "degc": "℃", "degf": "℉", "celsius": "℃",
    "fahrenheit": "℉", "angstrom": "Å", "micro": "µ", "permil": "‰",
    "pertenk": "‱",
}

GEOMETRY = {
    "angle": "∠", "measuredangle": "∡", "sphericalangle": "∢", "rightangle": "∟",
    "triangle": "△", "vartriangle": "△", "triangledown": "▽",
    "blacktriangle": "▲", "blacktriangledown": "▼",
    "square": "□", "box": "□", "rect": "▭", "rectangle": "▭",
    "parallelogram": "▱", "diamond": "⋄", "lozenge": "◊", "blacklozenge": "⧫",
    "circle": "○", "diameter": "⌀", "pentagon": "⬠", "hexagon": "⬡",
}

DOTS = {
    "ldots": "…", "dots": "…", "cdots": "⋯", "vdots": "⋮", "ddots": "⋱",
    "adots": "⋰", "iddots": "⋰", "rddots": "⋰", "dotsc": "…", "dotsb": "⋯",
    "dotsm": "⋯", "dotsi": "⋯", "dotso": "…", "cdotp": "·", "colon": ":",
}

DELIMITERS = {
    "langle": "⟨", "rangle": "⟩", "lAngle": "⟪", "rAngle": "⟫", "Rangle": "⟫",
    "llbracket": "⟦", "rrbracket": "⟧", "lbbrack": "⟦", "Rbrack": "⟧",
    "lfloor": "⌊", "rfloor": "⌋", "lceil": "⌈", "rceil": "⌉",
    "lvert": "|", "rvert": "|", "lVert": "‖", "rVert": "‖", "vert": "|",
    "Vert": "‖", "vbar": "│", "lbrace": "{", "rbrace": "}", "lbrack": "[",
    "rbrack": "]", "lparen": "(", "rparen": ")",
    "ulcorner": "⌜", "urcorner": "⌝", "llcorner": "⌞", "lrcorner": "⌟",
    "lgroup": "⟮", "rgroup": "⟯", "lmoustache": "⎰", "rmoustache": "⎱",
    "lmoust": "⎰", "rmoust": "⎱", "backslash": "\\",
    "overbrace": "⏞", "underbrace": "⏟", "overparen": "⏜", "underparen": "⏝",
    "overbracket": "⎴", "underbracket": "⎵", "overshell": "⏠", "undershell": "⏡",
}

SPACES = {
    "quad": "\u2003", "qquad": "\u2003\u2003", "enspace": "\u2002",
    "thinspace": "\u2009", "medspace": "\u205f", "thickspace": "\u2004",
    "emsp": "\u2003", "ensp": "\u2002", "thinsp": "\u2009", "medsp": "\u205f",
    "thicksp": "\u2004", "hairsp": "\u200a", "nbsp": "\u00a0", "zwsp": "\u200b",
    "zwnj": "\u200c", "funcapply": "\u2061", "itimes": "\u2062",
    "invisibletimes": "\u2062",
}

MISC = {
    "clubsuit": "♣", "diamondsuit": "♢", "heartsuit": "♡", "spadesuit": "♠",
    "flat": "♭", "natural": "♮", "sharp": "♯", "bigstar": "★", "maltese": "✠",
    "S": "§", "P": "¶", "copyright": "©", "pounds": "£", "euro": "€",
}

# Upright operator names: \sin -> sin (thin space added before a following letter).
FUNCTIONS = [
    "sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan",
    "arccot", "arcsec", "arccsc", "sinh", "cosh", "tanh", "coth", "sech", "csch",
    "arsinh", "arcosh", "artanh", "arcsinh", "arccosh", "arctanh",
    "log", "ln", "lg", "exp", "lim", "liminf", "limsup", "max", "min", "sup",
    "inf", "det", "dim", "ker", "deg", "gcd", "lcm", "arg", "argmax", "argmin",
    "hom", "Pr", "sgn", "rank", "tr", "Tr", "trace", "diag", "span", "erf",
    "erfc", "Var", "Cov", "mod", "bmod",
]

# Accents (combining marks). LaTeX form \hat{x} and Office postfix form x\hat both work.
ACCENTS = {
    "hat": "\u0302", "widehat": "\u0302", "check": "\u030c", "widecheck": "\u030c",
    "tilde": "\u0303", "widetilde": "\u0303", "acute": "\u0301", "grave": "\u0300",
    "dot": "\u0307", "ddot": "\u0308", "dddot": "\u20db", "ddddot": "\u20dc",
    "breve": "\u0306", "bar": "\u0304", "Bar": "\u033f", "mathring": "\u030a",
    "ring": "\u030a", "vec": "\u20d7", "hvec": "\u20d1", "rhvec": "\u20d1",
    "lhvec": "\u20d0", "lvec": "\u20d6", "tvec": "\u20e1",
    "overline": "\u0305", "overbar": "\u0305", "underline": "\u0332",
    "underbar": "\u0332", "ubar": "\u0332", "Ubar": "\u0333", "utilde": "\u0330",
    "overrightarrow": "\u20d7", "overleftarrow": "\u20d6",
    "overleftrightarrow": "\u20e1", "underrightarrow": "\u20ef",
    "underleftarrow": "\u20ee", "underleftrightarrow": "\u034d",
    "overrightharpoon": "\u20d1", "overleftharpoon": "\u20d0",
    "cancel": "\u0338", "sout": "\u0336", "strike": "\u0336",
    # Geometry: segment AB, ray AB, line AB, arc AB
    "seg": "\u0305", "ray": "\u20d7", "line": "\u20e1",
    "arc": "\u0311", "overarc": "\u0311", "wideparen": "\u0311",
}
SPAN_LR_ACCENTS = ["overleftrightarrow", "line", "tvec"]   # A⃖B⃗  (one arrow per end)
ARC_ACCENTS = ["arc", "overarc", "wideparen"]               # A͡B   (one arc spanning)

# ASCII shorthands, converted inside an expression. Longest match wins, so
# <=> never becomes ≤>, and <-> never becomes ←>.
ASCII_OPERATORS = {
    "<==>": "⟺", "<-->": "⟷", "==>": "⟹", "<==": "⟸", "-->": "⟶", "<--": "⟵",
    "|->": "↦", "<->": "↔", "<=>": "⇔", "===": "≡", "...": "…",
    "->": "→", "<-": "←", "=>": "⇒", "<=": "≤", ">=": "≥", "!=": "≠", "/=": "≠",
    "~=": "≈", "~~": "≈", "+-": "±", "-+": "∓", ":=": "≔", "=:": "≕",
    "<<": "≪", ">>": "≫", "/<": "≮", "/>": "≯",
}

VULGAR_FRACTIONS = {
    "1/2": "½", "1/3": "⅓", "2/3": "⅔", "1/4": "¼", "3/4": "¾", "1/5": "⅕",
    "2/5": "⅖", "3/5": "⅗", "4/5": "⅘", "1/6": "⅙", "5/6": "⅚", "1/7": "⅐",
    "1/8": "⅛", "3/8": "⅜", "5/8": "⅝", "7/8": "⅞", "1/9": "⅑", "1/10": "⅒",
    "0/3": "↉",
}

# Commands with special parsing (arguments). Handled inside the engine.
SPECIAL_COMMANDS = {
    "frac": ["frac", "dfrac", "tfrac", "cfrac"],
    "binom": ["binom", "dbinom", "tbinom"],
    "root": ["sqrt", "cbrt", "qdrt"],
    "over": ["overset", "underset", "stackrel"],
    "xarrow": ["xrightarrow", "xleftarrow", "xRightarrow", "xLeftarrow",
               "xleftrightarrow", "xLeftrightarrow", "xmapsto"],
    "pair": ["abs", "norm", "floor", "ceil", "ket", "bra", "braket", "set", "pmod"],
    "text": ["text", "textrm", "textnormal", "textup", "mathrm", "mbox", "hbox",
             "mathnormal", "operatorname"],
    "textfont": ["textbf", "textit", "textsf", "texttt"],
    "font": ["mathbf", "mathit", "mathcal", "mathscr", "mathfrak", "mathbb",
             "Bbb", "mathsf", "mathtt", "boldsymbol", "bm", "mathbfit",
             "mathsfbf", "mathbfsf"],
    "not": ["not"],
    "ignore": ["left", "right", "middle", "big", "Big", "bigg", "Bigg", "bigl",
               "bigr", "Bigl", "Bigr", "biggl", "biggr", "Biggl", "Biggr",
               "bigm", "Bigm", "biggm", "Biggm", "displaystyle", "textstyle",
               "scriptstyle", "scriptscriptstyle", "limits", "nolimits",
               "mathop", "mathrel", "mathbin", "mathord", "mathopen",
               "mathclose", "mathpunct", "mathinner", "nonumber", "notag"],
}
OFFICE_FONT_PREFIXES = {"double": "bb", "script": "cal", "fraktur": "frak"}  # \doubleR -> ℝ

SUPERSCRIPTS = {
    "0": "⁰", "1": "¹", "2": "²", "3": "³", "4": "⁴", "5": "⁵", "6": "⁶",
    "7": "⁷", "8": "⁸", "9": "⁹", "+": "⁺", "-": "⁻", "−": "⁻", "=": "⁼",
    "(": "⁽", ")": "⁾",
    "a": "ᵃ", "b": "ᵇ", "c": "ᶜ", "d": "ᵈ", "e": "ᵉ", "f": "ᶠ", "g": "ᵍ",
    "h": "ʰ", "i": "ⁱ", "j": "ʲ", "k": "ᵏ", "l": "ˡ", "m": "ᵐ", "n": "ⁿ",
    "o": "ᵒ", "p": "ᵖ", "r": "ʳ", "s": "ˢ", "t": "ᵗ", "u": "ᵘ", "v": "ᵛ",
    "w": "ʷ", "x": "ˣ", "y": "ʸ", "z": "ᶻ",
    "A": "ᴬ", "B": "ᴮ", "D": "ᴰ", "E": "ᴱ", "G": "ᴳ", "H": "ᴴ", "I": "ᴵ",
    "J": "ᴶ", "K": "ᴷ", "L": "ᴸ", "M": "ᴹ", "N": "ᴺ", "O": "ᴼ", "P": "ᴾ",
    "R": "ᴿ", "T": "ᵀ", "U": "ᵁ", "V": "ⱽ", "W": "ᵂ",
    "α": "ᵅ", "β": "ᵝ", "γ": "ᵞ", "δ": "ᵟ", "ε": "ᵋ", "ϵ": "ᵋ", "θ": "ᶿ",
    "ι": "ᶥ", "φ": "ᵠ", "ϕ": "ᵠ", "χ": "ᵡ",
    "∘": "°", "⊤": "ᵀ", "⊺": "ᵀ",          # 90^\circ -> 90°,  A^\top -> Aᵀ
}
# Already "raised" characters that may appear in a superscript unchanged.
SUPERSCRIPT_PASSTHROUGH = "°′″‴⁗†‡*∗"

SUBSCRIPTS = {
    "0": "₀", "1": "₁", "2": "₂", "3": "₃", "4": "₄", "5": "₅", "6": "₆",
    "7": "₇", "8": "₈", "9": "₉", "+": "₊", "-": "₋", "−": "₋", "=": "₌",
    "(": "₍", ")": "₎",
    "a": "ₐ", "e": "ₑ", "h": "ₕ", "i": "ᵢ", "j": "ⱼ", "k": "ₖ", "l": "ₗ",
    "m": "ₘ", "n": "ₙ", "o": "ₒ", "p": "ₚ", "r": "ᵣ", "s": "ₛ", "t": "ₜ",
    "u": "ᵤ", "v": "ᵥ", "x": "ₓ", "ə": "ₔ",
    "β": "ᵦ", "γ": "ᵧ", "ρ": "ᵨ", "φ": "ᵩ", "ϕ": "ᵩ", "χ": "ᵪ",
}

NEGATIONS = {  # \not X
    "=": "≠", "<": "≮", ">": "≯", "≤": "≰", "≥": "≱", "∈": "∉", "∋": "∌",
    "⊂": "⊄", "⊃": "⊅", "⊆": "⊈", "⊇": "⊉", "≡": "≢", "∼": "≁", "≃": "≄",
    "≅": "≇", "≈": "≉", "∃": "∄", "∣": "∤", "∥": "∦", "≺": "⊀", "≻": "⊁",
    "⊢": "⊬", "⊨": "⊭", "⊩": "⊮", "⊫": "⊯", "⊑": "⋢", "⊒": "⋣", "⊲": "⋪",
    "⊳": "⋫", "⊴": "⋬", "⊵": "⋭", "≍": "≭", "≲": "≴", "≳": "≵", "≶": "≸",
    "≷": "≹", "≼": "⋠", "≽": "⋡", "→": "↛", "←": "↚", "↔": "↮", "⇒": "⇏",
    "⇐": "⇍", "⇔": "⇎",
}

ESCAPES = {  # \<char>
    ",": "\u2009", ":": "\u205f", ">": "\u205f", ";": "\u2004", "!": "",
    "{": "{", "}": "}", "|": "‖", "_": "_", "^": "^", "%": "%", "#": "#",
    "&": "&", "$": "$", "/": "", "-": "",
}

# A fraction is written with parentheses when a part contains one of these.
FRACTION_PAREN_CHARS = "+-−±∓×÷⋅·∗/⁄=≠<>≤≥≈≡∼≃≅∝→←↔⇒⇐⇔∪∩∖∧∨⊕⊗∘,;:∈∉⊂⊃⊆⊇\u2009 "

# Entries of LibreOffice's own AutoCorrect table that fire inside math tokens and
# corrupt them (LibreOffice rewrites the text before Espanso replaces it).
LIBREOFFICE_CONFLICTS = [
    ".*->.*", ".*<-.*", "<->", "-->", "<--", "<-->", ".*...",
    "1/2", "1/4", "3/4", ".*(C)", ".*(R)", ".*(tm)",
]

# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  ENGINE -- runs inside Espanso (python -I -S -c ...) on every SPACE      ║
# ║  that follows a math token. Pure ASCII, Python 3.6+, never raises.      ║
# ╚══════════════════════════════════════════════════════════════════════════╝

ENGINE_TEMPLATE = r"""
import os, sys, unicodedata, zlib, binascii
try:
    sys.stderr = open(os.devnull, 'w')
except Exception:
    pass
LIMIT = @LIMIT@
SQRT_STYLE = '@SQRT@'
MINUS = @MINUS@
THIN = @THIN@
def _ld(s):
    d = {}
    for e in s.split(','):
        if e:
            k, v = e.split(':')
            d[k] = ''.join([chr(int(h, 16)) for h in v.split('+') if h])
    return d
def _lc(s):
    d = {}
    for k, v in _ld(s).items():
        d[chr(int(k, 16))] = v
    return d
SYM = _ld(zlib.decompress(binascii.a2b_base64('@SYMZ@')).decode('ascii'))
BIG = set('@BIG@'.split(','))
FUN = set('@FUN@'.split(','))
ACC = _ld('@ACC@')
OPS = {}
for _k, _v in _ld('@OPS@').items():
    OPS[bytes.fromhex(_k).decode('ascii')] = _v
VUL = {}
for _k, _v in _ld('@VUL@').items():
    VUL[bytes.fromhex(_k).decode('ascii')] = _v
SUP = _lc('@SUP@')
SUB = _lc('@SUB@')
NEG = _lc('@NEG@')
ESC = _lc('@ESC@')
SUPOK = set(SUP.values()) | set(_ld('x:@SUPPASS@')['x'])
SUBOK = set(SUB.values())
OPCH = set(_ld('x:@OPCH@')['x'])
SPANLR = set('@SPANLR@'.split(','))
ARC = set('@ARC@'.split(','))
SPECIAL = @SPECIAL@
KIND = {}
for _g, _names in SPECIAL.items():
    for _n in _names:
        KIND[_n] = _g
OFFICE = @OFFICE@
FONT = {
    'bf': (0x1D400, 0x1D41A, 0x1D7CE, {}),
    'it': (0x1D434, 0x1D44E, 0, {'h': 0x210E}),
    'bi': (0x1D468, 0x1D482, 0, {}),
    'cal': (0x1D49C, 0x1D4B6, 0, {'B': 0x212C, 'E': 0x2130, 'F': 0x2131, 'H': 0x210B,
            'I': 0x2110, 'L': 0x2112, 'M': 0x2133, 'R': 0x211B, 'e': 0x212F,
            'g': 0x210A, 'o': 0x2134}),
    'frak': (0x1D504, 0x1D51E, 0, {'C': 0x212D, 'H': 0x210C, 'I': 0x2111,
             'R': 0x211C, 'Z': 0x2128}),
    'bb': (0x1D538, 0x1D552, 0x1D7D8, {'C': 0x2102, 'H': 0x210D, 'N': 0x2115,
           'P': 0x2119, 'Q': 0x211A, 'R': 0x211D, 'Z': 0x2124}),
    'sf': (0x1D5A0, 0x1D5BA, 0x1D7E2, {}),
    'sfbf': (0x1D5D4, 0x1D5EE, 0x1D7EC, {}),
    'tt': (0x1D670, 0x1D68A, 0x1D7F6, {}),
}
FONTCMD = {'mathbf': 'bf', 'textbf': 'bf', 'boldsymbol': 'bi', 'bm': 'bi',
           'mathbfit': 'bi', 'mathit': 'it', 'textit': 'it', 'mathcal': 'cal',
           'mathscr': 'cal', 'mathfrak': 'frak', 'mathbb': 'bb', 'Bbb': 'bb',
           'mathsf': 'sf', 'textsf': 'sf', 'mathsfbf': 'sfbf', 'mathbfsf': 'sfbf',
           'mathtt': 'tt', 'texttt': 'tt'}
XAR = {'xrightarrow': '\u2192', 'xleftarrow': '\u2190', 'xRightarrow': '\u21d2',
       'xLeftarrow': '\u21d0', 'xleftrightarrow': '\u2194',
       'xLeftrightarrow': '\u21d4', 'xmapsto': '\u21a6'}
PAIR = {'abs': ('|', '|', None), 'norm': ('\u2016', '\u2016', '\u2016'),
        'floor': ('\u230a', '\u230b', None), 'ceil': ('\u2308', '\u2309', None),
        'ket': ('|', '\u27e9', '\u27e9'), 'bra': ('\u27e8', '|', '\u27e8'),
        'braket': ('\u27e8', '\u27e9', None), 'set': ('{', '}', None),
        'pmod': (' (mod ', ')', None)}
OVEREQ = {'def': '\u225d', '?': '\u225f', '\u25b3': '\u225c', '\u2206': '\u225c',
          '\u0394': '\u225c', '\u2227': '\u2259', '\u2228': '\u225a',
          '\u22c6': '\u225b', 'm': '\u225e', '\u2218': '\u2257', '.': '\u2250',
          '\u22c5': '\u2250'}
DELIMSIZE = set(['left', 'right', 'middle', 'big', 'Big', 'bigg', 'Bigg', 'bigl',
                 'bigr', 'Bigl', 'Bigr', 'biggl', 'biggr', 'Biggl', 'Biggr',
                 'bigm', 'Bigm', 'biggm', 'Biggm'])
PRIMES = ['\u2032', '\u2033', '\u2034', '\u2057']
LET = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'
DIG = '0123456789'
APOS = chr(39)
DIFF = 'd\u2202\u03b4\u0394\u2206\u2207'

def isin(c, s):
    return c != '' and c in s

def known(name):
    if name in SYM or name in FUN or name in ACC or name in KIND:
        return True
    for p in OFFICE:
        if name.startswith(p) and len(name) == len(p) + 1 and name[-1] in LET:
            return True
    return False

def font(style, s):
    up, lo, dg, ex = FONT[style]
    r = []
    for ch in s:
        if ch in ex:
            r.append(chr(ex[ch]))
        elif 'A' <= ch <= 'Z':
            r.append(chr(up + ord(ch) - 65))
        elif 'a' <= ch <= 'z':
            r.append(chr(lo + ord(ch) - 97))
        elif dg and '0' <= ch <= '9':
            r.append(chr(dg + ord(ch) - 48))
        else:
            r.append(ch)
    return ''.join(r)

def clusters(s):
    out = []
    for ch in s:
        if out and unicodedata.combining(ch):
            out[-1] += ch
        else:
            out.append(ch)
    return out

def nbase(s):
    return len([c for c in s if not unicodedata.combining(c)])

def balanced(s):
    d = 0
    for c in s:
        if c == '(':
            d += 1
        elif c == ')':
            d -= 1
            if d < 0:
                return False
    return d == 0

def wrapped(s):
    return len(s) >= 2 and s[0] == '(' and s[-1] == ')' and balanced(s[1:-1])

def wrap(s):
    if nbase(s) <= 1 or wrapped(s):
        return s
    return '(' + s + ')'

def fwrap(s):
    if s and not wrapped(s):
        for c in s:
            if c in OPCH:
                return '(' + s + ')'
    return s

def script_map(s, table, ok):
    r = []
    for ch in s:
        if ch == '\u2009':
            continue
        if unicodedata.combining(ch):
            r.append(ch)
        elif ch in table:
            r.append(table[ch])
        elif ch in ok:
            r.append(ch)
        else:
            return None
    return ''.join(r)

def sup(s):
    return script_map(s, SUP, SUPOK)

def sub(s):
    return script_map(s, SUB, SUBOK)

def mark_each(s, m):
    return ''.join([c + m for c in clusters(s.replace('\u2009', '')) if not c.isspace()])

def span_lr(s):
    cs = [c for c in clusters(s.replace('\u2009', '')) if not c.isspace()]
    if len(cs) <= 1:
        return ''.join([c + '\u20e1' for c in cs])
    mid = ''.join([c + '\u0305' for c in cs[1:-1]])
    return cs[0] + '\u20d6' + mid + cs[-1] + '\u20d7'

def arc(s):
    cs = [c for c in clusters(s.replace('\u2009', '')) if not c.isspace()]
    if len(cs) <= 1:
        return ''.join([c + '\u0311' for c in cs])
    return '\u0361'.join(cs)

def isatom(a):
    if a and all([c in DIG + '.' for c in a]):
        return True
    return nbase(a) <= 1

def radical(idx):
    if idx in (None, '', '2'):
        return '\u221a'
    if idx == '3':
        return '\u221b'
    if idx == '4':
        return '\u221c'
    s = sup(idx)
    return (s if s is not None else idx) + '\u221a'

def rootbody(a):
    if isatom(a):
        return a
    if SQRT_STYLE == 'parens':
        return a if wrapped(a) else '(' + a + ')'
    if wrapped(a):
        a = a[1:-1]
    return mark_each(a, '\u0305')

def frac(a, b):
    if a and b and all([c in DIG for c in a + b]):
        k = a + '/' + b
        if k in VUL:
            return VUL[k]
        return sup(a) + '\u2044' + sub(b)
    den = fwrap(b)
    if den == b and nbase(b) > 1 and not all([c in DIG for c in b]) and b[0] not in DIFF:
        den = '(' + b + ')'
    return fwrap(a) + '/' + den

def render(items):
    out = []
    for k in range(len(items)):
        t, kind = items[k]
        out.append(t)
        if THIN and kind in ('func', 'big') and k + 1 < len(items):
            nx = items[k + 1][0]
            if nx and unicodedata.category(nx[0])[0] in 'LN':
                out.append('\u2009')
    return ''.join(out)

class Cv(object):
    def __init__(self, t, frac_ok=True):
        self.t = t
        self.n = len(t)
        self.i = 0
        self.frac_ok = frac_ok

    def pk(self, k=0):
        j = self.i + k
        return self.t[j] if 0 <= j < self.n else ''

    def close(self, j, o, c):
        d = 0
        t = self.t
        while j < self.n:
            ch = t[j]
            if ch == '\\':
                j += 2
                continue
            if ch == o:
                d += 1
            elif ch == c:
                d -= 1
                if d == 0:
                    return j
            j += 1
        return -1

    def sub(self, s, frac_ok=None):
        return Cv(s, self.frac_ok if frac_ok is None else frac_ok).run()

    def run(self):
        return render(self.seq())

    def seq(self):
        items = []
        while self.i < self.n:
            c = self.t[self.i]
            if c == '^' or c == '_':
                base = items.pop() if items else ['', 'none']
                items.append(self.script(base))
                continue
            it = self.atom()
            if it is None:
                continue
            if it[1] == 'post':
                if items:
                    items[-1] = [items[-1][0] + it[0], items[-1][1]]
                else:
                    items.append([it[0], 'other'])
                continue
            items.append(it)
        return items

    def atom(self):
        t = self.t
        i = self.i
        c = t[i]
        if c == '\\':
            return self.command()
        if c == '{':
            j = self.close(i, '{', '}')
            if j < 0:
                self.i = i + 1
                return ['{', 'other']
            self.i = j + 1
            return [self.sub(t[i + 1:j]), 'group']
        if c in LET:
            j = i
            while j < self.n and t[j] in LET:
                j += 1
            self.i = j
            w = t[i:j]
            return [w, 'word' if len(w) > 1 else 'letter']
        if c in DIG:
            j = i
            while j < self.n and t[j] in DIG:
                j += 1
            if self.frac_ok and j + 1 < self.n and t[j] == '/' and t[j + 1] in DIG:
                k = j + 1
                while k < self.n and t[k] in DIG:
                    k += 1
                key = t[i:k]
                prev = t[i - 1] if i > 0 else ''
                nxt = t[k] if k < self.n else ''
                if key in VUL and not isin(prev, DIG + '/.,') and not isin(nxt, DIG + '/.,'):
                    self.i = k
                    return [VUL[key], 'num']
            self.i = j
            return [t[i:j], 'num']
        for L in (4, 3, 2):
            s = t[i:i + L]
            if len(s) == L and s in OPS:
                self.i = i + L
                return [OPS[s], 'op']
        if c == APOS:
            j = i
            while j < self.n and t[j] == APOS:
                j += 1
            self.i = j
            k = j - i
            out = ''
            while k > 0:
                out += PRIMES[min(k, 4) - 1]
                k -= min(k, 4)
            return [out, 'other']
        self.i = i + 1
        if c == '-' and MINUS:
            return ['\u2212', 'op']
        if c == '~':
            return [' ', 'space']
        return [c, 'other']

    def command(self):
        t = self.t
        i = self.i
        j = i + 1
        while j < self.n and t[j] in LET:
            j += 1
        name = t[i + 1:j]
        if not name:
            if j >= self.n:
                self.i = j
                return ['\\', 'other']
            ch = t[j]
            self.i = j + 1
            if ch in ESC:
                v = ESC[ch]
                if not v:
                    return None
                return [v, 'space' if v.isspace() else 'other']
            return ['\\' + ch, 'other']
        if not known(name) and len(name) > 1 and known(name[:-1]):
            j -= 1
            name = name[:-1]
        self.i = j
        r = self.dispatch(name)
        if r is False:
            self.i = j
            return ['\\' + name, 'other']
        return r

    def dispatch(self, name):
        g = KIND.get(name)
        if g == 'text' or g == 'textfont':
            s = self.graw()
            if s is None:
                return False
            if name in FONTCMD:
                s = font(FONTCMD[name], s)
            return [s, 'func' if name == 'operatorname' else 'text']
        if g == 'font':
            c = self.pk()
            if c == '{':
                a = self.garg()
            elif c and (c in LET or c in DIG):
                a = c
                self.i += 1
            else:
                a = None
            if a is None:
                return False
            return [font(FONTCMD[name], a), 'sym']
        if g == 'frac':
            a = self.arg1()
            b = self.arg1() if a is not None else None
            if b is None:
                return False
            return [frac(a, b), 'frac']
        if g == 'binom':
            a = self.arg1()
            b = self.arg1() if a is not None else None
            if b is None:
                return False
            p = sup(a)
            q = sub(b)
            if p is not None and q is not None:
                return [p + 'C' + q, 'sym']
            return ['C(' + a + ',' + b + ')', 'sym']
        if g == 'root':
            idx = None
            if name == 'sqrt' and self.pk() == '[':
                k = self.close(self.i, '[', ']')
                if k > 0:
                    idx = self.sub(self.t[self.i + 1:k], False)
                    self.i = k + 1
            elif name == 'cbrt':
                idx = '3'
            elif name == 'qdrt':
                idx = '4'
            a = self.arg1(True, True)
            if a is None:
                return [radical(idx), 'sym']
            return [radical(idx) + rootbody(a), 'sym']
        if g == 'over':
            a = self.arg1()
            b = self.arg1() if a is not None else None
            if b is None:
                return False
            if name == 'underset':
                s = sub(a)
                return [b + (s if s is not None else '_' + wrap(a)), 'sym']
            if b == '=' and a in OVEREQ:
                return [OVEREQ[a], 'op']
            s = sup(a)
            return [b + (s if s is not None else '^' + wrap(a)), 'sym']
        if g == 'xarrow':
            below = None
            if self.pk() == '[':
                k = self.close(self.i, '[', ']')
                if k > 0:
                    below = self.sub(self.t[self.i + 1:k])
                    self.i = k + 1
            above = self.garg()
            r = XAR[name]
            if above:
                s = sup(above)
                r += s if s is not None else '^' + wrap(above)
            if below:
                s = sub(below)
                r += s if s is not None else '_' + wrap(below)
            return [r, 'op']
        if g == 'pair':
            l, r, bare = PAIR[name]
            if bare is not None:
                a = self.garg()
                if a is None:
                    return [bare, 'sym']
            else:
                a = self.arg1()
                if a is None:
                    return False
            return [l + a + r, 'group']
        if g == 'not':
            save = self.i
            it = self.atom() if self.i < self.n else None
            if it is None or not it[0] or it[1] == 'post':
                self.i = save
                return ['\u0338', 'post']
            s = it[0]
            if s in NEG:
                return [NEG[s], 'op']
            cs = clusters(s)
            return [cs[0] + '\u0338' + ''.join(cs[1:]), 'op']
        if g == 'ignore':
            if name in DELIMSIZE and self.pk() == '.':
                self.i += 1
            return None
        if name in ACC:
            m = ACC[name]
            c = self.pk()
            a = None
            if c == '{':
                a = self.garg()
            elif c == '\\' and isin(self.pk(1), LET):
                a = self.arg1()
            elif c and (c in DIG or ord(c) > 127):
                a = self.arg1()
            if a is None:
                return [m, 'post']
            if name in SPANLR:
                return [span_lr(a), 'sym']
            if name in ARC:
                return [arc(a), 'sym']
            return [mark_each(a, m), 'sym']
        for p in OFFICE:
            if name.startswith(p) and len(name) == len(p) + 1 and name[-1] in LET:
                return [font(OFFICE[p], name[-1]), 'sym']
        if name in FUN:
            return ['mod' if name == 'bmod' else name, 'func']
        if name in SYM:
            return [SYM[name], 'big' if name in BIG else 'sym']
        return False

    def graw(self):
        if self.pk() == '{':
            j = self.close(self.i, '{', '}')
            if j < 0:
                return None
            s = self.t[self.i + 1:j]
            self.i = j + 1
            return s
        return None

    def garg(self):
        s = self.graw()
        return None if s is None else self.sub(s)

    def arg1(self, digits=False, paren=False):
        c = self.pk()
        if not c or c in '}^_':
            return None
        if c == '{':
            return self.garg()
        if paren and c == '(':
            j = self.close(self.i, '(', ')')
            if j < 0:
                return None
            s = self.t[self.i + 1:j]
            self.i = j + 1
            return self.sub(s)
        if c == '\\':
            save = self.i
            it = self.command()
            if it is None or it[1] == 'post':
                self.i = save
                return None
            return it[0]
        if digits and c in DIG:
            j = self.i
            while j < self.n and (self.t[j] in DIG or (self.t[j] == '.' and j + 1 < self.n
                                                       and self.t[j + 1] in DIG)):
                j += 1
            s = self.t[self.i:j]
            self.i = j
            return s
        self.i += 1
        if c == '-' and MINUS:
            return '\u2212'
        if c == APOS:
            return '\u2032'
        return c

    def script(self, base):
        op = self.t[self.i]
        self.i += 1
        start = self.i
        c = self.pk()
        S = None
        braced = False
        if c == '{' or c == '(':
            S = self.arg1(False, True)
            braced = S is not None
        elif c == '\\':
            S = self.arg1()
        else:
            sign = ''
            if isin(c, '+-') and (isin(self.pk(1), DIG) or isin(self.pk(1), LET)):
                sign = '\u2212' if (c == '-' and MINUS) else c
                self.i += 1
                c = self.pk()
            if isin(c, DIG):
                j = self.i
                while j < self.n and (self.t[j] in DIG or (self.t[j] == '.' and j + 1 < self.n
                                                           and self.t[j + 1] in DIG)):
                    j += 1
                S = sign + self.t[self.i:j]
                self.i = j
            elif isin(c, LET):
                if not isin(self.pk(1), LET):
                    S = sign + c
                    self.i += 1
            elif not sign and isin(c, '+-=*' + APOS):
                S = {'-': '\u2212' if MINUS else '-', APOS: '\u2032'}.get(c, c)
                self.i += 1
            elif not sign and c and ord(c) > 127 and not c.isspace():
                S = c
                self.i += 1
        if S is None:
            self.i = start
            return [base[0] + op, base[1]]
        raw = self.t[start:self.i]
        if op == '_' and not braced and base[1] == 'word':
            return [base[0] + op + raw, base[1]]
        conv = sup(S) if op == '^' else sub(S)
        if conv is None:
            conv = op + wrap(S)
        return [base[0] + conv, base[1]]

def convert(tok):
    if not tok:
        return tok
    low = tok.lower()
    if '://' in tok or low.startswith('www.') or '@' in tok:
        return tok
    if len(tok) > 2 and tok[0] in LET and tok[1] == ':' and tok[2] in '\\/':
        j = 3
        while j < len(tok) and tok[j] in LET:
            j += 1
        name = tok[3:j]
        if tok[2] == '/' or not (known(name) or (len(name) > 1 and known(name[:-1]))):
            return tok
    return unicodedata.normalize('NFC', Cv(tok).run())

def main():
    tok = os.environ.get('ESPANSO_TOK', '')
    out = tok
    try:
        if len(tok.encode('utf-8')) <= LIMIT:
            out = convert(tok)
    except Exception:
        out = tok
    try:
        sys.stdout.buffer.write(out.encode('utf-8'))
        sys.stdout.flush()
    except Exception:
        pass

if __name__ == '__main__':
    main()
"""

# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  GENERATOR                                                               ║
# ╚══════════════════════════════════════════════════════════════════════════╝

HEADER_MARK = "PowerPoint-Style Math AutoCorrect"
GENERATED_BY = "generated by generate_rules.py"
DEFAULT_BUFFER = 30            # Espanso's default max_regex_buffer_size
SAFETY_MARGIN = 5              # bytes kept free so a truncated buffer is never converted
MAX_COMMAND_LINE = 30000       # stay under Windows' 32,767-character limit

SYMBOL_TABLES = [
    ("Greek", GREEK), ("Hebrew", HEBREW), ("Binary operators", BINARY_OPERATORS),
    ("Relations", RELATIONS), ("Sets & number systems", SETS), ("Logic", LOGIC),
    ("Arrows", ARROWS), ("Big operators (take limits)", BIG_OPERATORS),
    ("Calculus & letter-like", CALCULUS_AND_LETTERLIKE), ("Geometry", GEOMETRY),
    ("Dots", DOTS), ("Delimiters", DELIMITERS), ("Spaces", SPACES), ("Misc", MISC),
]

NAME_RE = re.compile(r"^[A-Za-z]+$")
OP_RE = re.compile(r"^[!-/:-@\[-`{-~]{2,4}$")      # printable ASCII, no letters/digits


class GeneratorError(Exception):
    pass


def _clean_name(name):
    name = name[1:] if name.startswith("\\") else name
    if not NAME_RE.match(name):
        raise GeneratorError(f"Invalid command name {name!r}: use letters A-Z/a-z only.")
    return name


def build_tables():
    """Merge built-in and user tables, validating everything."""
    special = {n for names in SPECIAL_COMMANDS.values() for n in names}
    symbols = {}
    for title, table in SYMBOL_TABLES:
        for name, value in table.items():
            if name in symbols and symbols[name] != value:
                raise GeneratorError(f"\\{name} defined twice with different values ({title}).")
            symbols[name] = value
    functions = list(dict.fromkeys(FUNCTIONS))
    for name, value in USER_SYMBOLS.items():
        name = _clean_name(name)
        if not isinstance(value, str) or not value:
            raise GeneratorError(f"USER_SYMBOLS[\\{name}] must be a non-empty string.")
        if name in special or name in ACCENTS:
            raise GeneratorError(f"\\{name} is a built-in structural command; pick another name.")
        if name in functions:
            functions.remove(name)
        symbols[name] = value
    for name in USER_FUNCTIONS:
        name = _clean_name(name)
        symbols.pop(name, None)
        if name not in functions:
            functions.append(name)
    operators = dict(ASCII_OPERATORS)
    for seq, value in USER_OPERATORS.items():
        if not OP_RE.match(seq):
            raise GeneratorError(f"USER_OPERATORS key {seq!r}: 2-4 ASCII punctuation characters only.")
        if not isinstance(value, str) or not value:
            raise GeneratorError(f"USER_OPERATORS[{seq!r}] must be a non-empty string.")
        operators[seq] = value

    clash = (set(symbols) | set(functions)) & (special | set(ACCENTS))
    if clash:
        raise GeneratorError("Names used both as symbol and structural command: " + ", ".join(sorted(clash)))
    clash = set(symbols) & set(functions)
    if clash:
        raise GeneratorError("Names used both as symbol and function: " + ", ".join(sorted(clash)))
    for name in list(symbols) + functions:
        for prefix in OFFICE_FONT_PREFIXES:
            if name.startswith(prefix) and len(name) == len(prefix) + 1:
                raise GeneratorError(f"\\{name} collides with the Office font command \\{prefix}X.")
    for value in list(symbols.values()) + list(operators.values()):
        if any(ord(c) < 32 for c in value):
            raise GeneratorError(f"Replacement {value!r} contains control characters.")
    for table, label in ((SUPERSCRIPTS, "SUPERSCRIPTS"), (SUBSCRIPTS, "SUBSCRIPTS")):
        for k, v in table.items():
            if len(k) != 1 or len(v) != 1:
                raise GeneratorError(f"{label} entries must map one character to one character.")
    return {
        "symbols": symbols,
        "functions": functions,
        "operators": operators,
        "big": [n for n in BIG_OPERATORS if n in symbols],
        "special": special,
    }


def all_command_names(tables):
    names = set(tables["symbols"]) | set(tables["functions"]) | set(ACCENTS) | tables["special"]
    for prefix in OFFICE_FONT_PREFIXES:
        for ch in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz":
            names.add(prefix + ch)
    return names


# ---- engine ----------------------------------------------------------------

def _hexs(s):
    return "+".join(f"{ord(c):x}" for c in s)


def _enc_names(d):
    return ",".join(f"{k}:{_hexs(v)}" for k, v in sorted(d.items()))


def _enc_hexkeys(d):
    return ",".join(f"{k.encode('ascii').hex()}:{_hexs(v)}" for k, v in sorted(d.items()))


def _enc_chars(d):
    return ",".join(f"{ord(k):x}:{_hexs(v)}" for k, v in sorted(d.items()))


def buffer_limit(buffer_size):
    return max(8, int(buffer_size) - SAFETY_MARGIN)


def _compact(src):
    """Drop blank lines and use one space per indent level (keeps it readable, halves size)."""
    out = []
    for line in src.splitlines():
        if line.strip():
            n = len(line) - len(line.lstrip(" "))
            out.append(" " * (n // 4) + line.lstrip(" "))
    return "\n".join(out) + "\n"


def build_engine(tables, buffer_size=DEFAULT_BUFFER):
    special = {k: sorted(v) for k, v in SPECIAL_COMMANDS.items()}
    repl = {
        "@LIMIT@": str(buffer_limit(buffer_size)),
        "@SQRT@": "parens" if SETTINGS.get("sqrt_style") == "parens" else "vinculum",
        "@MINUS@": "True" if SETTINGS.get("minus_sign", True) else "False",
        "@THIN@": "True" if SETTINGS.get("thin_space_after_operators", True) else "False",
        "@SYMZ@": base64.b64encode(zlib.compress(
            _enc_names(tables["symbols"]).encode("ascii"), 9)).decode("ascii"),
        "@BIG@": ",".join(sorted(tables["big"])),
        "@FUN@": ",".join(sorted(tables["functions"])),
        "@ACC@": _enc_names(ACCENTS),
        "@OPS@": _enc_hexkeys(tables["operators"]),
        "@VUL@": _enc_hexkeys(VULGAR_FRACTIONS),
        "@SUP@": _enc_chars(SUPERSCRIPTS),
        "@SUB@": _enc_chars(SUBSCRIPTS),
        "@NEG@": _enc_chars(NEGATIONS),
        "@ESC@": _enc_chars(ESCAPES),
        "@SUPPASS@": _hexs(SUPERSCRIPT_PASSTHROUGH),
        "@OPCH@": _hexs(FRACTION_PAREN_CHARS),
        "@SPANLR@": ",".join(SPAN_LR_ACCENTS),
        "@ARC@": ",".join(ARC_ACCENTS),
        "@SPECIAL@": repr(special),
        "@OFFICE@": repr(OFFICE_FONT_PREFIXES),
    }
    src = ENGINE_TEMPLATE
    for key, value in repl.items():
        src = src.replace(key, value)
    src = _compact(src)
    leftover = re.findall(r"@[A-Z]+@", src)
    if leftover:
        raise GeneratorError(f"Unfilled engine placeholders: {leftover}")
    try:
        src.encode("ascii")
    except UnicodeEncodeError as exc:
        raise GeneratorError(f"Engine source must be pure ASCII: {exc}")
    compile(src, "<ppt-math-engine>", "exec")
    return src


def load_engine(src):
    """Load the engine in-process (used by --self-test and --try)."""
    ns = {"__name__": "ppt_math_engine"}
    exec(compile(src, "<ppt-math-engine>", "exec"), ns)
    return ns


# ---- regex -------------------------------------------------------------------

# Characters that can never be part of a math token (whitespace, control chars).
_EXCL = ("\\x00-\\x20\\x7f-\\xa0"
         "\u1680\u2000-\u200b\u2028\u2029\u202f\u205f\u3000\ufeff")
_RUST_META = set("\\.+*?()|[]{}^$#&-~")


def _rx_escape(s):
    """Escape for Rust regex 1.5 (Espanso): only meta characters may be escaped."""
    return "".join("\\" + c if c in _RUST_META else c for c in s)


def build_regex(tables):
    """One regex: a whitespace-free token that contains a math marker, then a space.

    The match is never anchored to its left context, so Espanso's leftmost match
    is always the whole token: nothing outside the token is deleted and the
    character before it (space, Enter, Tab ...) is never touched.
    """
    T = f"[^{_EXCL}]"
    NL = f"[^A-Za-z{_EXCL}]"
    tail = f"(?:{NL}{T}*)?"            # token ends here, or continues with a non-letter
    names = sorted(all_command_names(tables), key=lambda s: (-len(s), s))
    ops = sorted(tables["operators"], key=lambda s: (-len(s), s))
    alternatives = [
        "\\\\(?:" + "|".join(names) + ")[A-Za-z]?" + tail,           # \command
        "\\\\[,:;!>|{}_^%#$/\\-\\&]" + T + "*",                          # \, \{ \| ...
        "[\\^_][+\\-]?[A-Za-z]" + tail,                                  # x^n  x_i
        "[\\^_](?:[{(\\\\0-9+\\-=*']|[^\\x00-\\x7f" + _EXCL + "])" + T + "*",  # x^2 x^{..}
        "(?:" + "|".join(_rx_escape(o) for o in ops) + ")" + T + "*",   # -> <= ...
        "(?:" + "|".join(sorted(VULGAR_FRACTIONS)) + ")" + T + "*",      # 1/2
    ]
    return f"(?P<tok>{T}*(?:" + "|".join(alternatives) + "))\\x20$"


# ---- YAML ---------------------------------------------------------------------

def yaml_quote(s):
    """Double-quoted YAML scalar, pure ASCII (non-ASCII as \\u escapes)."""
    out = []
    for c in s:
        o = ord(c)
        if c == "\\":
            out.append("\\\\")
        elif c == '"':
            out.append('\\"')
        elif 0x20 <= o < 0x7F:
            out.append(c)
        elif o <= 0xFFFF:
            out.append(f"\\u{o:04x}")
        else:
            out.append(f"\\U{o:08x}")
    return '"' + "".join(out) + '"'


def render_match_yaml(regex, engine_src, python_cmd, inject_mode="auto", note=""):
    lines = [
        f"# {HEADER_MARK} for Espanso -- {GENERATED_BY}. Do not edit by hand:",
        "# edit scripts/generate_rules.py (USER_SYMBOLS etc.) and run it again.",
        "#",
        "# Type an expression WITHOUT spaces (LaTeX / PowerPoint linear format),",
        "# then press SPACE.  Examples: \\alpha  x^2  \\int_0^1  \\sum_{i=1}^{n}  a<=b",
    ]
    if note:
        lines.append(f"# {note}")
    lines += [
        "",
        "matches:",
        '  - label: "Math AutoCorrect (LaTeX / PowerPoint linear format)"',
        f"    regex: {yaml_quote(regex)}",
        '    replace: "{{out}} "',
    ]
    if inject_mode in ("clipboard", "keys"):
        lines.append(f"    force_mode: {inject_mode}")
    lines += [
        "    vars:",
        "      - name: out",
        "        type: script",
        "        inject_vars: false",
        "        params:",
        "          trim: false",
        "          args:",
        f"            - {yaml_quote(python_cmd)}",
        '            - "-I"',
        '            - "-S"',
        '            - "-c"',
        "            - |",
    ]
    for line in engine_src.splitlines():
        lines.append(("              " + line) if line.strip() else "")
    return "\n".join(lines) + "\n"


def render_libreoffice_config(match_filename):
    return (
        f"# {HEADER_MARK} -- {GENERATED_BY}.\n"
        "# Activates the math rules only while LibreOffice (soffice) is focused.\n"
        "# Note: app-specific configs need X11, Windows or macOS (not Wayland).\n"
        'filter_exec: "(?i)soffice"\n'
        "extra_includes:\n"
        f'  - "../match/{match_filename}"\n'
    )


def check_command_line(python_cmd, engine_src):
    length = len(python_cmd) + len(engine_src) + engine_src.count('"') + 64
    if length > MAX_COMMAND_LINE:
        raise GeneratorError(
            f"Engine is {length} characters; Windows allows ~32k on a command line. "
            "Remove some USER_SYMBOLS or shorten them.")


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  INSTALLER                                                               ║
# ╚══════════════════════════════════════════════════════════════════════════╝

GLOBAL_MATCH = Path("match") / "ppt_math.yml"
LO_MATCH = Path("match") / "_ppt_math.yml"
LO_CONFIG = Path("config") / "libreoffice_math.yml"


def espanso_executable():
    return shutil.which("espanso") or shutil.which("espansod")


def espanso_base_dir(override=None):
    if override:
        return Path(override).expanduser().resolve()
    exe = espanso_executable()
    if exe:
        try:
            res = subprocess.run([exe, "path", "config"], stdout=subprocess.PIPE,
                                 stderr=subprocess.DEVNULL, text=True, timeout=20)
            out = res.stdout.strip().splitlines()
            if res.returncode == 0 and out and Path(out[-1].strip()).is_dir():
                return Path(out[-1].strip())
        except (OSError, subprocess.SubprocessError):
            pass
    home = Path.home()
    if sys.platform == "darwin":
        candidates = [home / "Library" / "Application Support" / "espanso",
                      home / "Library" / "Preferences" / "espanso"]
    elif os.name == "nt":
        candidates = [Path(os.environ.get("APPDATA", home / "AppData" / "Roaming")) / "espanso"]
    else:
        xdg = os.environ.get("XDG_CONFIG_HOME")
        candidates = [Path(xdg) / "espanso" if xdg else home / ".config" / "espanso",
                      home / ".var" / "app" / "org.espanso.Espanso" / "config" / "espanso"]
    for c in candidates:
        if (c / "config").is_dir() or (c / "match").is_dir():
            return c
    return candidates[0]


def default_python(for_install):
    if not for_install:
        return "python3"
    if os.name == "nt":
        return sys.executable
    if sys.platform == "darwin":                  # stable across Homebrew upgrades
        found = shutil.which("python3")
        return found if found and found != "/usr/bin/python3" else sys.executable
    for cand in ("/usr/bin/python3", "/usr/local/bin/python3"):
        if os.access(cand, os.X_OK):
            return cand
    return shutil.which("python3") or sys.executable


def read_buffer_size(base):
    default_yml = base / "config" / "default.yml"
    try:
        text = default_yml.read_text(encoding="utf-8")
    except OSError:
        return DEFAULT_BUFFER
    m = re.search(r"^max_regex_buffer_size\s*:\s*(\d+)\s*(#.*)?$", text, re.M)
    return int(m.group(1)) if m else DEFAULT_BUFFER


def write_buffer_size(base, size):
    default_yml = base / "config" / "default.yml"
    default_yml.parent.mkdir(parents=True, exist_ok=True)
    text = default_yml.read_text(encoding="utf-8") if default_yml.exists() else ""
    if default_yml.exists():
        shutil.copy2(default_yml, default_yml.with_name("default.yml.bak"))
    line = f"max_regex_buffer_size: {size}"
    if re.search(r"^max_regex_buffer_size\s*:.*$", text, re.M):
        text = re.sub(r"^max_regex_buffer_size\s*:.*$", line, text, flags=re.M)
    else:
        text = text.rstrip("\n") + ("\n" if text else "") + line + "\n"
    atomic_write(default_yml, text)


def atomic_write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=".tmp-", suffix=".yml", dir=str(path.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as f:
            f.write(text)
        os.replace(tmp, path)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def is_ours(path):
    try:
        head = path.read_text(encoding="utf-8", errors="replace")[:4000]
    except OSError:
        return False
    return HEADER_MARK.lower() in head.lower() or "generate_rules.py" in head


def remove_if_ours(path, removed):
    if path.exists() and is_ours(path):
        path.unlink()
        removed.append(path)


def verify_python(python_cmd, engine_src):
    """Run the engine exactly as Espanso will. Returns (ok, message)."""
    env = dict(os.environ, ESPANSO_TOK="\\alpha^2")
    try:
        res = subprocess.run([python_cmd, "-I", "-S", "-c", engine_src], env=env,
                             stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30)
    except (OSError, subprocess.SubprocessError) as exc:
        return False, f"cannot run {python_cmd!r}: {exc}"
    got = res.stdout.decode("utf-8", "replace")
    if res.returncode != 0 or res.stderr.strip() or got != "α²":
        return False, (f"{python_cmd!r} returned code {res.returncode}, "
                       f"output {got!r}, errors {res.stderr.decode('utf-8', 'replace')[:300]!r}")
    return True, "ok"


# ---- diagnostics -------------------------------------------------------------

TRIGGER_LINE = re.compile(r"""^\s*-?\s*(trigger|regex)\s*:\s*(['"]?)(.*?)\2\s*(#.*)?$""")
LIST_ITEM = re.compile(r"""^\s*-\s*(['"]?)(.*?)\1\s*$""")


def _yaml_unquote(s, quote):
    if quote == "'":
        return s.replace("''", "'")
    if quote == '"':
        return re.sub(r'\\(["\\])', r"\1", s)
    return s


def scan_other_matches(base, own_paths, tables):
    """Find other copies of these rules and triggers that fire on the same keys."""
    copies, conflicts, regex_count = [], [], 0
    config_dir = base / "config"
    if config_dir.is_dir():
        for path in sorted(config_dir.glob("*.y*ml")):
            if path.resolve() != (base / LO_CONFIG).resolve() and is_ours(path):
                copies.append(path)
    match_dir = base / "match"
    if not match_dir.is_dir():
        return copies, conflicts, regex_count
    own = {p.resolve() for p in own_paths}
    risky_ops = sorted(tables["operators"], key=len, reverse=True)
    for path in sorted(list(match_dir.rglob("*.yml")) + list(match_dir.rglob("*.yaml"))):
        if path.resolve() in own:
            continue
        if is_ours(path):
            copies.append(path)
            continue
        try:
            lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
        except OSError:
            continue
        in_triggers = False
        for line in lines:
            m = TRIGGER_LINE.match(line)
            values = []
            if m:
                in_triggers = False
                if m.group(1) == "regex":
                    regex_count += 1
                    continue
                values.append((m.group(3), m.group(2)))
            elif re.match(r"^\s*-?\s*triggers\s*:\s*$", line):
                in_triggers = True
                continue
            elif in_triggers:
                lm = LIST_ITEM.match(line)
                if lm:
                    values.append((lm.group(2), lm.group(1)))
                else:
                    in_triggers = False
            for raw, quote in values:
                trig = _yaml_unquote(raw, quote)
                if ("\\" in trig or "^" in trig or "_" in trig
                        or any(op in trig for op in risky_ops)
                        or any(fr in trig for fr in VULGAR_FRACTIONS)):
                    conflicts.append((path, trig))
    return copies, conflicts, regex_count


def libreoffice_autocorr_dirs():
    home = Path.home()
    if os.name == "nt":
        roots = [Path(os.environ.get("APPDATA", home)) / "LibreOffice"]
    elif sys.platform == "darwin":
        roots = [home / "Library" / "Application Support" / "LibreOffice"]
    else:
        xdg = Path(os.environ.get("XDG_CONFIG_HOME", home / ".config"))
        roots = [xdg / "libreoffice",
                 home / ".var" / "app" / "org.libreoffice.LibreOffice" / "config" / "libreoffice",
                 home / "snap" / "libreoffice" / "current" / ".config" / "libreoffice"]
    return [r / "4" / "user" / "autocorr" for r in roots if (r / "4" / "user").is_dir()]


def scan_libreoffice_autocorrect():
    """Return (files_scanned, {file: [conflicting entries]}) for the user's own lists."""
    found, scanned = {}, []
    for d in libreoffice_autocorr_dirs():
        for dat in sorted(d.glob("acor_*.dat")):
            scanned.append(dat)
            try:
                with zipfile.ZipFile(dat) as z:
                    xml = z.read("DocumentList.xml").decode("utf-8", "replace")
            except (OSError, KeyError, zipfile.BadZipFile):
                continue
            names = {_xml_unescape(n) for n in re.findall(r'abbreviated-name="([^"]*)"', xml)}
            hits = [c for c in LIBREOFFICE_CONFLICTS if c in names]
            if hits:
                found[dat] = hits
    # A language without a user file still uses LibreOffice's built-in list.
    langs = {d.stem.lower() for d in scanned}
    complete = bool(scanned) and {"acor_en-gb", "acor_en-us"} <= langs
    return complete, found


def _xml_unescape(s):
    return (s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
             .replace("&apos;", "'").replace("&amp;", "&"))


LIBREOFFICE_ADVICE = """\
LibreOffice's own AutoCorrect rewrites some of these sequences *before* Espanso
replaces them, which corrupts the text (this is where "<-> becomes ←>" comes from).
In LibreOffice: Tools > AutoCorrect > AutoCorrect Options...
  * Replace tab: delete these entries (for each language you type in):
      {entries}
  * Options tab: untick "Automatic *bold*, /italic/, -strikeout- and _underline_"
    (and "Replace dashes" if you use --> or <--).
Or untick "Use replacement table" to switch LibreOffice's list off entirely."""


def print_diagnostics(base, own_paths, tables, buffer_size, lo_only, python_cmd, engine_src):
    print("\nDiagnostics")
    print("-----------")
    ok, msg = verify_python(python_cmd, engine_src)
    print(f"  Python engine ({python_cmd}): {'OK' if ok else 'FAILED - ' + msg}")
    limit = buffer_limit(buffer_size)
    print(f"  Regex buffer: {buffer_size} bytes -> expressions up to {limit} characters convert.")
    if buffer_size < 60:
        print("    Tip: --buffer-size 64 allows longer expressions (sets max_regex_buffer_size).")
    copies, conflicts, regex_count = scan_other_matches(base, own_paths, tables)
    for path in copies:
        pkg = path.parent.name if "packages" in path.parts else None
        hint = f" (run: espanso uninstall {pkg})" if pkg else " (delete it)"
        print(f"  ! Another copy of these rules is loaded: {path}{hint}")
    if conflicts:
        print("  ! Other Espanso triggers that may fire together with these rules:")
        for path, trig in conflicts[:25]:
            print(f"      {trig:30}  in {path}")
        if len(conflicts) > 25:
            print(f"      ... and {len(conflicts) - 25} more")
    if regex_count:
        print(f"  Note: {regex_count} other regex rule(s) found; make sure none end with a space.")
    if not copies and not conflicts:
        print("  No conflicting Espanso rules found.")
    if lo_only and os.environ.get("XDG_SESSION_TYPE", "").lower() == "wayland":
        print("  ! Wayland session: Espanso cannot detect the focused app, so the\n"
              "    LibreOffice-only filter never activates. Use a global install instead.")
    scanned, found = scan_libreoffice_autocorrect()
    for dat, hits in found.items():
        print(f"  ! LibreOffice AutoCorrect conflicts in {dat}: {', '.join(hits)}")
    if scanned and not found:
        print("  LibreOffice AutoCorrect: no conflicting entries in your profile.")
    else:
        if not scanned:
            print("  LibreOffice AutoCorrect: using the built-in lists, which contain conflicting entries.")
        print("\n" + LIBREOFFICE_ADVICE.format(entries="  ".join(LIBREOFFICE_CONFLICTS)))


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  SELF-TEST                                                               ║
# ╚══════════════════════════════════════════════════════════════════════════╝

TEST_CASES = [
    # --- symbols, chaining, prefix-safety ---
    ("\\alpha", "α"), ("\\in", "∈"), ("\\infty", "∞"), ("\\int", "∫"),
    ("\\inc", "∆"), ("\\to", "→"), ("\\top", "⊤"), ("\\alpha\\beta", "αβ"),
    ("\\Deltax", "Δx"), ("\\alpha,", "α,"), ("(\\alpha)", "(α)"),
    ("\\epsilon", "ϵ"), ("\\varepsilon", "ε"), ("\\phi", "ϕ"), ("\\varphi", "φ"),
    ("\\R^n", "ℝⁿ"), ("f:\\R\\to\\R", "f:ℝ→ℝ"), ("\\doubleR", "ℝ"),
    ("\\mathbb{N}", "ℕ"), ("\\mathcal{L}", "ℒ"), ("\\mathfrak{g}", "𝔤"),
    ("\\mathbf{x}", "𝐱"), ("\\mathbb{1}", "𝟙"),
    # --- ASCII operators: longest match always wins ---
    ("<->", "↔"), ("<-", "←"), ("->", "→"), ("<=>", "⇔"), ("<=", "≤"),
    ("=>", "⇒"), (">=", "≥"), ("!=", "≠"), ("<-->", "⟷"), ("-->", "⟶"),
    ("==>", "⟹"), ("<==>", "⟺"), ("|->", "↦"), ("a<->b", "a↔b"),
    ("x<=y", "x≤y"), ("+-", "±"), ("~=", "≈"), (":=", "≔"), ("...", "…"),
    # --- super/subscripts ---
    ("x^2", "x²"), ("x^10", "x¹⁰"), ("x^{n+1}", "xⁿ⁺¹"), ("x^(n+1)", "xⁿ⁺¹"),
    ("x^{-1}", "x⁻¹"), ("x^-1", "x⁻¹"), ("e^{i\\theta}", "eⁱᶿ"), ("x_1", "x₁"),
    ("a_{ij}", "aᵢⱼ"), ("x_n^2", "xₙ²"), ("\\alpha_1", "α₁"), ("A^T", "Aᵀ"),
    ("A^\\top", "Aᵀ"), ("90^\\circ", "90°"), ("f^\\prime", "f′"), ("f''(x)=x^2", "f″(x)=x²"),
    ("e^{i\\pi}", "e^(iπ)"), ("v_y", "v_y"), ("T_{max}", "Tₘₐₓ"), ("T_{cd}", "T_(cd)"),
    ("file_2", "file_2"), ("x_ab", "x_ab"), ("E=mc^2", "E=mc²"),
    ("e^{-x^2}", "e⁻ˣ²"), ("2^n", "2ⁿ"),
    # --- big operators & limits ---
    ("\\int_0^1", "∫₀¹"), ("\\int_a^b", "∫ₐᵇ"), ("\\int_{0}^{\\infty}", "∫₀^∞"),
    ("\\int_{-\\infty}^{\\infty}", "∫_(−∞)^∞"), ("\\sum_{i=1}^{n}", "∑ᵢ₌₁ⁿ"),
    ("\\sum_{i=1}^{n}x_i", "∑ᵢ₌₁ⁿ\u2009xᵢ"), ("\\prod_{k=1}^n", "∏ₖ₌₁ⁿ"),
    ("\\lim_{x\\to0}", "lim_(x→0)"), ("\\lim_{n\\to\\infty}", "lim_(n→∞)"),
    ("\\oint_C", "∮_C"), ("\\bigcup_{i=1}^n", "⋃ᵢ₌₁ⁿ"), ("\\sum\\limits_i", "∑ᵢ"),
    ("\\sin\\theta", "sin\u2009θ"), ("\\sin^2x", "sin²\u2009x"), ("\\sin(x)", "sin(x)"),
    ("\\max_x", "maxₓ"),
    # --- roots & fractions ---
    ("\\sqrt{x}", "√x"), ("\\sqrt{2}", "√2"), ("\\sqrt{x+1}", "√x̅+̅1̅"),
    ("\\sqrt(x+1)", "√x̅+̅1̅"), ("\\sqrt{x^2+y^2}", "√x̅²̅+̅y̅²̅"), ("\\sqrt[3]{x}", "∛x"),
    ("\\sqrt[n]{x}", "ⁿ√x"), ("\\sqrt", "√"), ("\\cbrt{8}", "∛8"), ("\\sqrtx", "√x"),
    ("\\frac{1}{2}", "½"), ("\\frac12", "½"), ("\\frac{3}{16}", "³⁄₁₆"),
    ("\\frac{a+b}{c}", "(a+b)/c"), ("\\frac{dy}{dx}", "dy/dx"), ("\\frac{1}{x}", "1/x"),
    ("\\frac{1}{2\\pi}", "1/(2π)"), ("\\frac{d^2y}{dx^2}", "d²y/dx²"), ("\\frac{\\pi}{4}", "π/4"),
    ("1/2", "½"), ("x+1/2", "x+½"), ("11/2", "11/2"), ("1/20", "1/20"),
    ("\\binom{n}{k}", "ⁿCₖ"),
    # --- sets & logic ---
    ("A\\cup{B}", "A∪B"), ("A\\cap{B}", "A∩B"), ("A\\setminus{B}", "A∖B"),
    ("A\\subseteq{B}", "A⊆B"), ("x\\notin{A}", "x∉A"), ("\\emptyset", "∅"),
    ("\\{x\\in\\R\\mid{x}>0\\}", "{x∈ℝ∣x>0}"), ("\\forall{x}\\exists{y}", "∀x∃y"),
    ("\\not\\in", "∉"), ("\\not=", "≠"), ("\\not\\subset", "⊄"), ("p\\land{q}", "p∧q"),
    ("\\neg{p}", "¬p"), ("A\\symdiff{B}", "A∆B"), ("\\powerset(A)", "𝒫(A)"),
    # --- geometry ---
    ("\\overline{AB}", "A̅B̅"), ("\\seg{AB}", "A̅B̅"), ("\\ray{AB}", "A⃗B⃗"),
    ("\\line{AB}", "A⃖B⃗"), ("\\overleftrightarrow{AB}", "A⃖B⃗"), ("\\arc{AB}", "A͡B"),
    ("\\angle{ABC}", "∠ABC"), ("\\triangle{ABC}", "△ABC"), ("AB\\parallel{CD}", "AB∥CD"),
    ("AB\\perp{CD}", "AB⊥CD"), ("\\measuredangle{A}", "∡A"),
    # --- accents ---
    ("\\vec{v}", "v⃗"), ("\\hat{x}", "x̂"), ("\\bar{x}", "x̄"), ("\\dot{x}", "ẋ"),
    ("\\ddot{x}", "ẍ"), ("\\tilde{n}", "ñ"), ("x\\hat", "x̂"), ("\\vec\\alpha", "α⃗"),
    # --- delimiters & misc ---
    ("\\abs{x}", "|x|"), ("\\norm{v}", "‖v‖"), ("\\floor{x}", "⌊x⌋"), ("\\ceil{x}", "⌈x⌉"),
    ("\\ket{\\psi}", "|ψ⟩"), ("\\bra{\\phi}", "⟨ϕ|"), ("\\langle{x}\\rangle", "⟨x⟩"),
    ("\\left(x\\right)", "(x)"), ("a\\equiv{b}\\pmod{n}", "a≡b (mod n)"),
    ("\\text{if}", "if"), ("x\\in[0,1]", "x∈[0,1]"), ("\\alpha~b", "α b"),
    ("\\overset{def}{=}", "≝"), ("\\xrightarrow{f}", "→ᶠ"), ("\\hbar", "ℏ"),
    ("\\partial_x{f}", "∂ₓf"), ("\\nabla\\cdot\\vec{E}", "∇⋅E⃗"), ("x^2-1", "x²−1"),
    # --- things that must stay untouched ---
    ("\\foo", "\\foo"), ("http://a.b/x_1", "http://a.b/x_1"),
    ("C:\\temp\\alpha", "C:\\temp\\alpha"), ("me@x.com^2", "me@x.com^2"),
    ("D:/x_1", "D:/x_1"), ("T:\\R^n\\to\\R^m", "T:ℝⁿ→ℝᵐ"), ("x=\\not\\equiv", "x=≢"),
]

# Typed text -> must NOT trigger Espanso at all (no flicker, no change).
NON_MATH_TOKENS = [
    "hello", "don't", "snake_case", "__init__", "e-mail", "24/7", "and/or",
    "C:\\Users\\Pavan", "\\newline", "x_axis", "@user", "#tag", "50%",
    "(see", "note:", "A-level", "re-use", "$5", "a.b.c", "x", "ab", "\\",
]


def run_self_test(tables, engine_src, regex, python_cmd=None):
    eng = load_engine(engine_src)
    convert = eng["convert"]
    failures = 0
    for src, expected in TEST_CASES:
        try:
            got = convert(src)
        except Exception as exc:          # the real engine never raises; tests should see why
            got = f"<exception {exc!r}>"
        if got != unicodedata.normalize("NFC", expected):
            failures += 1
            print(f"  FAIL  {src!r:34} -> {got!r}   (expected {expected!r})")
    # Every test input must be recognised as a math token by the regex.
    rx = re.compile(regex.replace("\\x20$", "\\x20\\Z"))
    for src, expected in TEST_CASES:
        if expected == src:
            continue
        m = rx.search(src + " ")
        if not m or m.group("tok") != src:
            failures += 1
            print(f"  FAIL  regex does not capture {src!r} (got {m.group('tok') if m else None!r})")
    for word in NON_MATH_TOKENS:
        if rx.search(word + " "):
            failures += 1
            print(f"  FAIL  regex fires on ordinary text {word!r}")
    # Space-separated tokens are independent; leading context is never consumed.
    for typed, tok in [("foo \\alpha ", "\\alpha"), ("a\r\\beta ", "\\beta"),
                       ("x^2 y^2 ", "y^2")]:
        m = rx.search(typed)
        if not m or m.group("tok") != tok:
            failures += 1
            print(f"  FAIL  {typed!r} captured {m.group('tok') if m else None!r}, expected {tok!r}")
    # The generated YAML must parse back to exactly the regex and engine.
    yaml_text = render_match_yaml(regex, engine_src, python_cmd or "python3")
    try:
        import yaml  # optional
        doc = yaml.safe_load(yaml_text)
        m0 = doc["matches"][0]
        args = m0["vars"][0]["params"]["args"]
        if m0["regex"] != regex or args[-1] != engine_src or m0["replace"] != "{{out}} ":
            failures += 1
            print("  FAIL  YAML round-trip mismatch")
    except ImportError:
        print("  (PyYAML not installed: YAML round-trip check skipped)")
    if python_cmd:
        ok, msg = verify_python(python_cmd, engine_src)
        if not ok:
            failures += 1
            print(f"  FAIL  engine via {python_cmd}: {msg}")
    total = 2 * len(TEST_CASES) + len(NON_MATH_TOKENS) + 4
    print(f"Self-test: {total - failures}/{total} checks passed.")
    return failures == 0


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║  CLI                                                                     ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def repo_root():
    here = Path(__file__).resolve().parent
    return here.parent if here.name == "scripts" else here


def list_commands(tables, needle):
    rows = []
    for title, table in SYMBOL_TABLES:
        for name in table:
            if name in tables["symbols"]:
                rows.append((title, "\\" + name, tables["symbols"][name]))
    for name, value in USER_SYMBOLS.items():
        rows.append(("User", "\\" + _clean_name(name), value))
    for name in tables["functions"]:
        rows.append(("Functions", "\\" + name, name))
    for name, mark in ACCENTS.items():
        rows.append(("Accents", "\\" + name + "{x}", "x" + mark))
    for group, names in SPECIAL_COMMANDS.items():
        for name in names:
            rows.append(("Structures", "\\" + name, "(" + group + ")"))
    for seq, value in tables["operators"].items():
        rows.append(("ASCII", seq, value))
    for seq, value in VULGAR_FRACTIONS.items():
        rows.append(("Fractions", seq, value))
    for prefix, style in OFFICE_FONT_PREFIXES.items():
        rows.append(("Fonts", "\\" + prefix + "A..Z/a..z", "(" + style + ")"))
    seen = set()
    for title, cmd, value in rows:
        if needle and needle.lower() not in (cmd + " " + title).lower():
            continue
        if (title, cmd) in seen:
            continue
        seen.add((title, cmd))
        print(f"  {title:28} {cmd:26} {value}")


def main(argv=None):
    p = argparse.ArgumentParser(
        description="Generate / install PowerPoint-style Math AutoCorrect rules for Espanso.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Type an expression without spaces, then press SPACE to convert it.")
    p.add_argument("--install", action="store_true", help="install into your Espanso config")
    p.add_argument("--libreoffice-only", action="store_true",
                   help="activate the rules only inside LibreOffice")
    p.add_argument("--uninstall", action="store_true", help="remove the installed rules")
    p.add_argument("--check", action="store_true", help="only run the conflict diagnostics")
    p.add_argument("--self-test", action="store_true", help="run the built-in conversion tests")
    p.add_argument("--try", dest="try_exprs", nargs="+", metavar="EXPR",
                   help="show how expressions convert")
    p.add_argument("--list", dest="list_filter", nargs="?", const="", metavar="TEXT",
                   help="list every supported command (optionally filtered)")
    p.add_argument("--stdout", action="store_true", help="print the match YAML instead of writing it")
    p.add_argument("--output", type=Path, help="output path when not installing")
    p.add_argument("--espanso-dir", help="Espanso config directory (default: auto-detect)")
    p.add_argument("--python", help="Python interpreter Espanso should run")
    p.add_argument("--buffer-size", type=int, metavar="N",
                   help="set Espanso's max_regex_buffer_size (default 30) to allow longer expressions")
    p.add_argument("--inject-mode", choices=["auto", "clipboard", "keys"], default="auto",
                   help="how Espanso types the result (clipboard avoids app autocorrect re-firing)")
    args = p.parse_args(argv)

    try:
        tables = build_tables()
    except GeneratorError as exc:
        print(f"Error in symbol tables: {exc}", file=sys.stderr)
        return 2

    if args.list_filter is not None:
        list_commands(tables, args.list_filter)
        return 0

    installing = args.install or args.uninstall or args.check
    base = espanso_base_dir(args.espanso_dir) if installing else None
    if args.buffer_size is not None:
        if not 16 <= args.buffer_size <= 512:
            print("--buffer-size must be between 16 and 512.", file=sys.stderr)
            return 2
        buffer_size = args.buffer_size
    elif base is not None:
        buffer_size = read_buffer_size(base)
    else:
        buffer_size = DEFAULT_BUFFER
    python_cmd = args.python or default_python(args.install or args.check)

    try:
        engine_src = build_engine(tables, buffer_size)
        regex = build_regex(tables)
        check_command_line(python_cmd, engine_src)
    except GeneratorError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2

    if args.try_exprs:
        eng = load_engine(engine_src)
        for expr in args.try_exprs:
            print(f"  {expr}  ->  {eng['convert'](expr)}")
        return 0

    if args.self_test:
        ok = run_self_test(tables, engine_src, regex, args.python or (sys.executable if not args.install else python_cmd))
        return 0 if ok else 1

    if args.uninstall:
        removed = []
        for rel in (GLOBAL_MATCH, LO_MATCH, LO_CONFIG):
            remove_if_ours(base / rel, removed)
        for path in removed:
            print(f"Removed {path}")
        if not removed:
            print(f"Nothing to remove in {base}.")
        return 0

    if args.check:
        own = [base / GLOBAL_MATCH, base / LO_MATCH]
        lo_mode = args.libreoffice_only or is_ours(base / LO_CONFIG)
        print(f"Espanso config: {base}")
        for rel in (GLOBAL_MATCH, LO_MATCH, LO_CONFIG):
            state = "installed" if is_ours(base / rel) else "not installed"
            print(f"  {rel}: {state}")
        print_diagnostics(base, own, tables, buffer_size, lo_mode, python_cmd, engine_src)
        return 0

    if args.install:
        ok, msg = verify_python(python_cmd, engine_src)
        if not ok:
            print(f"Error: the engine does not run with {python_cmd}: {msg}\n"
                  "Pass a working interpreter with --python /path/to/python3.", file=sys.stderr)
            return 2
        if not (base / "config").is_dir() and not (base / "match").is_dir():
            print(f"Warning: {base} does not look like an Espanso config folder "
                  "(use --espanso-dir).", file=sys.stderr)
        if args.buffer_size is not None:
            write_buffer_size(base, buffer_size)
            print(f"Set max_regex_buffer_size: {buffer_size} in {base / 'config' / 'default.yml'}")
        removed = []
        if args.libreoffice_only:
            match_path = base / LO_MATCH
            atomic_write(match_path, render_match_yaml(
                regex, engine_src, python_cmd, args.inject_mode, "Active only in LibreOffice."))
            atomic_write(base / LO_CONFIG, render_libreoffice_config(LO_MATCH.name))
            remove_if_ours(base / GLOBAL_MATCH, removed)
            print(f"Installed {match_path}\nInstalled {base / LO_CONFIG}")
        else:
            match_path = base / GLOBAL_MATCH
            atomic_write(match_path, render_match_yaml(
                regex, engine_src, python_cmd, args.inject_mode))
            remove_if_ours(base / LO_MATCH, removed)
            remove_if_ours(base / LO_CONFIG, removed)
            print(f"Installed {match_path}")
        for path in removed:
            print(f"Removed old {path}")
        print(f"{len(all_command_names(tables))} commands, {len(tables['operators'])} ASCII operators, "
              f"{len(VULGAR_FRACTIONS)} fractions. Espanso reloads automatically.")
        print_diagnostics(base, [match_path], tables, buffer_size, args.libreoffice_only,
                          python_cmd, engine_src)
        return 0

    # Repository build (no install)
    if args.stdout:
        sys.stdout.write(render_match_yaml(regex, engine_src, python_cmd, args.inject_mode))
        return 0
    root = repo_root()
    if args.libreoffice_only:
        out_dir = args.output or (root / "dist" / "libreoffice")
        atomic_write(out_dir / LO_MATCH, render_match_yaml(
            regex, engine_src, python_cmd, args.inject_mode, "Active only in LibreOffice."))
        atomic_write(out_dir / LO_CONFIG, render_libreoffice_config(LO_MATCH.name))
        print(f"Wrote {out_dir / LO_MATCH}\nWrote {out_dir / LO_CONFIG}\n"
              "Copy the 'match' and 'config' folders into your Espanso config folder,\n"
              "or simply run with --install --libreoffice-only.")
    else:
        out = args.output or (root / "package.yml")
        atomic_write(out, render_match_yaml(regex, engine_src, python_cmd, args.inject_mode))
        print(f"Wrote {out} ({len(all_command_names(tables))} commands, "
              f"{len(tables['operators'])} ASCII operators, {len(VULGAR_FRACTIONS)} fractions).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
