"""Build RL_for_SWE_Agents.xlsx — one sheet per concept (Concept / Math / Example).

Source material: github.com/RajatDandekar/Mini-SWE-RL (incl. its swe-rl-ipr/ sub-project).
Every example is live Excel formulas: change a blue cell on a yellow background and the
sheet recomputes. Run:  python build_workbook.py
"""

import math
import re

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, ScatterChart, Series
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

OUT = "RL_for_SWE_Agents.xlsx"
FONT = "Arial"
NAVY = "1F3864"
FILL = {
    "concept": "DDEBF7",
    "math": "E4DFEC",
    "example": "E2EFDA",
    "mathline": "F4F1F8",
    "hdr": "D9E1F2",
    "input": "FFF2CC",
    "total": "F2F2F2",
}
FIRST, LAST = 2, 11  # content spans columns B..K
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def font(**kw):
    kw.setdefault("name", FONT)
    kw.setdefault("size", 10)
    return Font(**kw)


def fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)


class Sheet:
    """Tiny layout helper: writes top-to-bottom, tracks the current row, and resolves
    «name» placeholders in formulas to absolute cell addresses at the end."""

    def __init__(self, wb, name, title, idea, widths=None):
        self.ws = wb.create_sheet(name)
        self.name = name
        self.names = {}
        ws = self.ws
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 2
        ws.column_dimensions["B"].width = 24
        for c in range(3, 30):
            ws.column_dimensions[L(c)].width = 13
        for letter, w in (widths or {}).items():
            ws.column_dimensions[letter].width = w
        link = ws.cell(1, 2, "← Back to Index")
        link.hyperlink = "#'Index'!A1"
        link.font = font(color="0563C1", underline="single", size=9)
        self.r = 2
        self._merged(title, font(bold=True, size=16, color="FFFFFF"), fill(NAVY), height=30)
        self._merged("In one line: " + idea, font(italic=True, size=11, color=NAVY), height=None)
        self.r += 1

    # -- low-level -------------------------------------------------------------
    def _merged(self, text, fnt, fll=None, height=None, c1=FIRST, c2=LAST, indent=0):
        ws, r = self.ws, self.r
        if c2 > c1:
            ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        cell = ws.cell(r, c1, text)
        cell.font = fnt
        cell.alignment = Alignment(wrap_text=True, vertical="center", indent=indent)
        if fll is not None:
            cell.fill = fll
        if height is None:
            width_units = sum(ws.column_dimensions[L(c)].width for c in range(c1, c2 + 1))
            cpl = max(20, width_units * (1.05 if fnt.sz <= 10 else 0.9))
            lines = sum(max(1, math.ceil(len(seg) / cpl)) for seg in str(text).split("\n"))
            height = max(16, 14 * lines + 4)
        ws.row_dimensions[r].height = height
        self.r += 1
        return cell

    # -- blocks ----------------------------------------------------------------
    def section(self, label, kind):
        self.r += 1
        self._merged(label, font(bold=True, size=12, color=NAVY), fill(FILL[kind]), height=22)

    def sub(self, text):
        self._merged(text, font(bold=True, size=10.5, color=NAVY), height=18)

    def para(self, text, italic=False, color="000000", bold=False):
        self._merged(text, font(italic=italic, color=color, bold=bold))

    def bullets(self, items):
        for t in items:
            self._merged("•  " + t, font(), indent=1)

    def note(self, text):
        self._merged(text, font(italic=True, color="595959"))

    def math(self, lines):
        for t in lines:
            self._merged(t, font(size=11), fill(FILL["mathline"]), indent=2)

    def legend(self, rows):
        ws, r = self.ws, self.r
        for c, h in ((FIRST, "Symbol"), (FIRST + 1, "Meaning")):
            cell = ws.cell(r, c, h)
            cell.font = font(bold=True)
            cell.fill = fill(FILL["hdr"])
        ws.merge_cells(start_row=r, start_column=FIRST + 1, end_row=r, end_column=LAST)
        self.r += 1
        for sym, meaning in rows:
            ws.cell(self.r, FIRST, sym).font = font(bold=True)
            ws.cell(self.r, FIRST).alignment = Alignment(vertical="center")
            self._merged(meaning, font(), c1=FIRST + 1)

    def gap(self, n=1):
        self.r += n

    def kv(self, rows, value_col=FIRST + 1):
        """rows: (name, label, value, number_format, note). Constants are inputs (blue on yellow)."""
        ws = self.ws
        for name, label, value, fmt, note_ in rows:
            r = self.r
            lc = ws.cell(r, FIRST, label)
            lc.font = font(bold=True)
            lc.alignment = Alignment(wrap_text=True, vertical="center")
            vc = ws.cell(r, value_col, value)
            self._style_value(vc, value, fmt)
            vc.border = BOX
            h = 18
            if note_:
                ws.merge_cells(start_row=r, start_column=value_col + 1, end_row=r, end_column=LAST)
                nc = ws.cell(r, value_col + 1, note_)
                nc.font = font(italic=True, color="595959", size=9)
                nc.alignment = Alignment(wrap_text=True, vertical="center")
                cpl = (LAST - value_col) * 13 * 1.15
                h = max(h, 13 * math.ceil(len(note_) / cpl) + 4)
            h = max(h, 13 * math.ceil(len(label) / 26) + 4)
            ws.row_dimensions[r].height = h
            if name:
                self.names[name] = f"${L(value_col)}${r}"
            self.r += 1

    def table(self, headers, nrows, row_fn, fmts=None, label_cols=(0,), merge_last=False,
              row_height=None, header_height=None, bold_rows=()):
        """row_fn(i, r) -> list of cell values for data row i (0-based) at sheet row r.
        Numeric constants outside label_cols are styled as inputs. Returns (first, last)."""
        ws = self.ws
        fmts = fmts or {}
        hr = self.r
        ncols = len(headers)
        for j, h in enumerate(headers):
            c = ws.cell(hr, FIRST + j, h)
            c.font = font(bold=True, size=9.5)
            c.fill = fill(FILL["hdr"])
            c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            c.border = BOX
        if merge_last and FIRST + ncols - 1 < LAST:
            ws.merge_cells(start_row=hr, start_column=FIRST + ncols - 1, end_row=hr, end_column=LAST)
        longest = max(len(str(h)) for h in headers)
        ws.row_dimensions[hr].height = header_height or (16 if longest < 14 else 30 if longest < 30 else 44)
        first = hr + 1
        for i in range(nrows):
            r = first + i
            vals = row_fn(i, r)
            for j, v in enumerate(vals):
                c = ws.cell(r, FIRST + j, v)
                if j in label_cols or isinstance(v, str) and not v.startswith("="):
                    c.font = font(bold=(i in bold_rows))
                    c.alignment = Alignment(wrap_text=True, vertical="center",
                                            horizontal="center" if isinstance(v, (int, float)) else None)
                    if isinstance(v, (int, float)) and j in fmts:
                        c.number_format = fmts[j]
                else:
                    self._style_value(c, v, fmts.get(j))
                    if i in bold_rows:
                        c.font = font(bold=True, color=c.font.color)
                if i in bold_rows:
                    c.fill = fill(FILL["total"])
                c.border = BOX
            if merge_last and FIRST + ncols - 1 < LAST:
                ws.merge_cells(start_row=r, start_column=FIRST + ncols - 1, end_row=r, end_column=LAST)
            if row_height:
                ws.row_dimensions[r].height = row_height
        self.r = first + nrows
        return first, first + nrows - 1

    def _style_value(self, cell, value, fmt):
        is_input = isinstance(value, (int, float)) and not isinstance(value, bool)
        if is_input:
            cell.font = font(color="0000FF")
            cell.fill = fill(FILL["input"])
        else:
            cell.font = font()
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if fmt:
            cell.number_format = fmt

    def resolve(self):
        pat = re.compile(r"«([A-Za-z0-9_]+)»")
        for row in self.ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and "«" in c.value:
                    c.value = pat.sub(lambda m: self.names[m.group(1)], c.value)


def col(i):
    """Letter of the i-th content column (0 = B)."""
    return L(FIRST + i)


def rng(i, a, b):
    return f"${col(i)}${a}:${col(i)}${b}"


def line_chart(title, x_title, y_title, width=16, height=8.5):
    ch = LineChart()
    ch.title = title
    ch.x_axis.title = x_title
    ch.y_axis.title = y_title
    ch.width, ch.height = width, height
    ch.style = 2
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    return ch


# =============================================================================
SHEETS = []  # (sheet name, part, one-line idea)


def new(wb, name, part, title, idea, widths=None):
    SHEETS.append((name, part, title, idea))
    return Sheet(wb, name, title, idea, widths)


# -----------------------------------------------------------------------------
def s01(wb):
    s = new(wb, "01 RL Setup", "Foundations", "The RL setup for a coding agent",
            "The model tries fixes, a test suite scores them, and training makes high-scoring fixes more likely.",
            widths={"C": 36})
    s.section("CONCEPT", "concept")
    s.bullets([
        "Reinforcement learning (RL) teaches by trial and score, not by copying answers. The model writes a fix, "
        "something checks it, and the model's weights are nudged toward whatever scored well.",
        "State / observation = what the model sees: the bug description + the buggy function.",
        "Action = the text the model writes (the proposed fix).",
        "Reward = the score for that action: did the tests pass?",
        "Policy π_θ = the language model itself; θ = its weights. The policy is a probability distribution over "
        "every text it could write.",
        "Episode = one attempt from reset() to reward. Rollout = a recorded episode (prompt, response, reward).",
        "Why RL instead of supervised fine-tuning (SFT)? SFT needs the correct answer for every example. RL only "
        "needs a way to CHECK an answer. Tests are a cheap, automatic checker, which is why this is called RL "
        "with verifiable rewards.",
        "Analogy: practising past exam papers with an answer checker but no worked solutions. You learn which "
        "approaches get marked right.",
    ])
    s.sub("Where it lives in the repo")
    s.para("env.py → CodeFixEnv.reset() / step() / get_prompt();  agent.py → run_agent_on_puzzle();  "
           "production equivalent: R2E-Gym (Docker repos) + DeepSWE agent.")

    s.section("MATH", "math")
    s.math([
        "π_θ(y | x) = probability that the model with weights θ writes response y for prompt x",
        "R(x, y) ∈ {0, 1} = reward (1 if every test passes)",
        "J(θ) = E_x E_{y ~ π_θ(·|x)} [ R(x, y) ]   → the expected reward, i.e. the solve rate",
        "For a single prompt:  J = Σ_y π_θ(y | x) · R(x, y)   (probability × reward, summed over possible answers)",
        "Training goal: find θ that maximises J(θ).",
    ])
    s.legend([
        ("x", "the prompt (bug report + buggy code)"),
        ("y", "a full response the model could write"),
        ("θ", "the model's weights, the only thing training changes"),
        ("E[·]", "expected value = probability-weighted average"),
    ])

    s.section("EXAMPLE: expected reward of a toy policy on one puzzle (sum_range off-by-one)", "example")
    s.note("Pretend the model can only write 4 different answers. Edit the blue probabilities (each column "
           "must sum to 1) and rewards to see J change.")
    rows = [
        ("A", "range(a, b + 1)   ← correct", 0.35, 1, 0.55),
        ("B", "range(a, b)   ← unchanged bug", 0.30, 0, 0.15),
        ("C", "range(a + 1, b + 1)   ← new bug", 0.20, 0, 0.10),
        ("D", "(a + b) * (b - a + 1) // 2   ← correct, different idea", 0.15, 1, 0.20),
    ]
    a, b = s.r + 1, s.r + 4
    s.table(
        ["Answer", "What the model writes", "P(answer) now", "Reward R", "P × R", "P(answer) after training",
         "P × R after"],
        5,
        lambda i, r: (list(rows[i][:4]) + [f"=D{r}*E{r}", rows[i][4], f"=G{r}*E{r}"]) if i < 4 else
        ["Total", "J = Σ P × R  (expected reward)", f"=SUM(D{a}:D{b})", "", f"=SUM(F{a}:F{b})",
         f"=SUM(G{a}:G{b})", f"=SUM(H{a}:H{b})"],
        fmts={2: "0.00", 3: "0", 4: "0.00", 5: "0.00", 6: "0.00"},
        label_cols=(0, 1), bold_rows=(4,),
    )
    s.note("What to notice: J now = 0.50 and after = 0.75. Training never sees the 'right answer'. It only moves "
           "probability toward rows whose reward is 1. The next sheets show how.")
    return s


def s02(wb):
    s = new(wb, "02 Env & Reward", "Foundations", "Environment and binary reward",
            "The environment hands out a bug, runs the tests on the proposed fix, and returns 1 only if every test passes.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Environment = the thing that gives tasks and scores answers. Mini-SWE-RL's CodeFixEnv holds 45 puzzles "
        "(15 easy, 15 medium, 15 hard). Each puzzle is a description + a buggy function + a list of assert tests.",
        "reset(puzzle) returns the observation. step(code) executes the fix, runs every test, and returns the reward.",
        "Binary reward = 1 only if ALL tests pass (same as DeepSWE). It is hard to game partially, but sparse: a fix "
        "passing 4 of 5 tests scores the same 0 as garbage.",
        "Partial reward = fraction of tests passed gives a denser signal, but invites 'pass the easy tests and stop'. "
        "Meta's SWE-RL used a continuous reward instead (difflib similarity to the reference patch).",
        "Generated code is untrusted, so it needs isolation and a timeout. In the repo, the buggy fix_binary_search "
        "loops forever and env.step has no timeout. swe-rl-ipr runs candidates in a subprocess with a timeout "
        "(runner.py) and in gVisor sandboxes on Modal at scale.",
        "The grader IS the specification. env.step re-runs the fix before every test, so state cannot leak between "
        "tests. Because of that, the buggy code for hard_default_mutable and med_shallow_copy_accumulate already "
        "passes all their tests. The trainer's own scorer runs the fix once for all tests and does catch them. Two "
        "graders means two different specs.",
    ])
    s.sub("Where it lives in the repo")
    s.para("env.py CodeFixEnv.step();  puzzles.py / puzzles_medium.py / puzzles_hard.py;  "
           "grpo_trainer_v2.py collect_rollouts_ollama() (its own inline scorer);  swe-rl-ipr/ipr/runner.py run_cases().")

    s.section("MATH", "math")
    s.math([
        "passed = Σ_j 1[test_j passes]          (1[·] is 1 when the condition is true, else 0)",
        "Binary reward:   R = 1 if passed = N, else 0",
        "Partial reward:  R = passed / N",
    ])
    s.legend([("N", "number of tests for the puzzle"), ("1[·]", "indicator function")])

    s.section("EXAMPLE: five candidate fixes for is_balanced(s)  (1 = test passed, 0 = failed)", "example")
    data = [
        ("Correct fix", 1, 1, 1, 1, 1),
        ("Handles () only, ignores [] {}", 1, 1, 1, 1, 0),
        ("Off-by-one in the loop", 1, 0, 1, 0, 1),
        ("Syntax error", 0, 0, 0, 0, 0),
        ("Hardcoded test outputs", 1, 1, 1, 1, 1),
    ]
    s.table(["Candidate fix", "Test 1", "Test 2", "Test 3", "Test 4", "Test 5", "Passed", "Total N",
             "Binary reward", "Partial reward"], 5,
            lambda i, r: list(data[i]) + [f"=SUM(C{r}:G{r})", f"=COUNT(C{r}:G{r})", f"=IF(H{r}=I{r},1,0)",
                                           f"=H{r}/I{r}"],
            fmts={1: "0", 2: "0", 3: "0", 4: "0", 5: "0", 6: "0", 7: "0", 8: "0", 9: "0%"}, row_height=28)
    s.note("What to notice: binary reward gives 'Handles () only' (80% of tests) the same 0 as a syntax error. "
           "The hardcoded answer earns a full 1 even though it is wrong. That is reward hacking (sheet 17).")
    return s


def s03(wb):
    s = new(wb, "03 Temperature", "Foundations", "Sampling and temperature",
            "Temperature controls how random the model's choices are. GRPO needs randomness so attempts differ.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "At every step the model gives each possible next token a score (a logit). Softmax turns scores into "
        "probabilities, and one token is sampled.",
        "Temperature T divides the logits before softmax. T → 0: always the top token (greedy, used for "
        "evaluation). T = 1: the model's own distribution. T > 1: flatter, more random.",
        "GRPO needs diversity inside a group. If all 8 attempts are identical they get identical rewards, and there "
        "is nothing to compare. Mini-SWE-RL samples at T = 0.8 and evaluates at T = 0.",
        "Too high → more nonsense and a lower solve rate. Too low → no variety and no learning signal.",
    ])
    s.sub("Where it lives in the repo")
    s.para("agent.py query_ollama(temperature=…);  grpo_trainer_v2.py Config.temperature = 0.8;  "
           "evaluate_*() uses temperature 0 / do_sample=False.")
    s.section("MATH", "math")
    s.math([
        "p_i = exp(z_i / T) / Σ_j exp(z_j / T)",
        "Greedy decoding = pick argmax_i z_i (the limit T → 0)",
        "Entropy H = − Σ_i p_i · ln p_i   (0 = certain; higher = more random)",
        "P(all G samples pick the same token) = Σ_i p_i^G  ≈ p_top^G when one token dominates",
    ])
    s.legend([("z_i", "logit (raw score) of token i"), ("T", "temperature"), ("p_i", "probability of token i")])

    s.section("EXAMPLE: next token after  'for i in range(a, '  when fixing sum_range", "example")
    s.kv([("G", "Group size G", 8, "0", "used in the last row")])
    tr = s.r
    s.ws.cell(tr, FIRST, "Temperature T →").font = font(bold=True)
    temps = [0.2, 0.8, 1.0, 1.5]
    for j, t in enumerate(temps):
        c = s.ws.cell(tr, FIRST + 2 + j, t)
        s._style_value(c, t, "0.0")
        c.border = BOX
    s.r += 1
    toks = [("b + 1   ← correct", 3.0), ("b", 2.6), ("b - 1", 0.5), ("len(b)", 0.2), ("a", -1.0)]
    a, b = s.r + 1, s.r + 5

    def row(i, r):
        if i < 5:
            return [toks[i][0], toks[i][1]] + [
                f"=EXP($C{r}/{col(2 + j)}${tr})/SUMPRODUCT(EXP($C${a}:$C${b}/{col(2 + j)}${tr}))" for j in range(4)]
        if i == 5:
            return ["Sum (check = 1)", ""] + [f"=SUM({col(2 + j)}{a}:{col(2 + j)}{b})" for j in range(4)]
        if i == 6:
            return ["Entropy H", ""] + [f"=-SUMPRODUCT({col(2 + j)}{a}:{col(2 + j)}{b},LN({col(2 + j)}{a}:{col(2 + j)}{b}))"
                                         for j in range(4)]
        return ["P(all G samples identical)", ""] + [f"=SUMPRODUCT({col(2 + j)}{a}:{col(2 + j)}{b}^«G»)"
                                                      for j in range(4)]

    f1, f2 = s.table(["Next-token option", "Logit z", "p at T1", "p at T2", "p at T3", "p at T4"], 8, row,
                     fmts={1: "0.0", 2: "0.0%", 3: "0.0%", 4: "0.0%", 5: "0.0%"}, label_cols=(0,),
                     bold_rows=(5, 6, 7))
    for rr in (f2 - 1,):
        for j in range(4):
            s.ws.cell(rr, FIRST + 2 + j).number_format = "0.00"
    s.note("What to notice: at T = 0.2 almost every sample is 'b + 1', so all 8 attempts look alike and GRPO learns "
           "nothing. At T = 0.8 the wrong 'b' shows up often, giving a mix of right and wrong attempts to compare.")
    ch = BarChart()
    ch.type = "col"
    ch.title = "Probability of each token at each temperature"
    ch.y_axis.title = "probability"
    ch.width, ch.height = 16, 8
    ch.add_data(Reference(s.ws, min_col=FIRST + 2, max_col=FIRST + 5, min_row=f1 - 1, max_row=f1 + 4),
                titles_from_data=True)
    ch.set_categories(Reference(s.ws, min_col=FIRST, min_row=f1, max_row=f1 + 4))
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    s.ws.add_chart(ch, f"B{s.r + 1}")
    s.r += 18
    return s


def s04(wb):
    s = new(wb, "04 Log-Probs", "Foundations", "Log-probability of a response",
            "A response's probability is the product of its token probabilities. We add their logs instead, and only over tokens the model wrote.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "A response is a sequence of tokens. Its probability = P(token 1) × P(token 2 | token 1) × … (chain rule).",
        "Multiplying hundreds of numbers below 1 underflows to 0 on a computer, so we add logarithms instead. "
        "The log-probability is always ≤ 0; closer to 0 means more likely.",
        "Only response tokens count. Prompt tokens are masked (weight 0), because the model shouldn't be "
        "rewarded or punished for text it was given.",
        "Sum or mean? grpo_trainer.py (v1) sums token log-probs, so long answers get larger gradients. "
        "grpo_trainer_v2.py uses the mean, so every answer counts equally (see sheet 13).",
        "How it's computed: one forward pass over prompt + response gives logits at every position. The logits at "
        "position t predict token t+1 (shift by one). Take log_softmax, then pick the value for the token that "
        "actually appeared.",
    ])
    s.sub("Where it lives in the repo")
    s.para("grpo_trainer_v2.py compute_log_probs(): logits[prompt_len-1:-1] → log_softmax → gather(response_tokens).")
    s.section("MATH", "math")
    s.math([
        "log π(y | x) = Σ_t log π(y_t | x, y_<t)          (sum over response tokens only)",
        "π(y | x) = exp( log π(y | x) )",
        "mean log-prob = (1 / |y|) · Σ_t log π(y_t | x, y_<t)",
        "perplexity = exp( − mean log-prob )   (≈ 'how many tokens the model was choosing between')",
    ])
    s.legend([("y_t", "the t-th token of the response"), ("y_<t", "all tokens before position t"),
              ("|y|", "number of response tokens")])
    s.section("EXAMPLE: log-prob of the fix  'return a + b'", "example")
    toks = [("Fix", "prompt", 0.10), (" add", "prompt", 0.05), ("(a, b):", "prompt", 0.30),
            ("return", "response", 0.92), (" a", "response", 0.85), (" +", "response", 0.40),
            (" b", "response", 0.95), ("<eos>", "response", 0.90)]
    a, b = s.r + 1, s.r + len(toks)
    s.table(["Position", "Token", "Part of", "Mask (1 = trained)", "p(token | before)", "ln p",
             "Mask × ln p", "p used in product"], len(toks),
            lambda i, r: [i + 1, toks[i][0], toks[i][1], f'=IF(D{r}="response",1,0)', toks[i][2],
                          f"=LN(F{r})", f"=E{r}*G{r}", f"=IF(E{r}=1,F{r},1)"],
            fmts={0: "0", 4: "0.00", 5: "0.000", 6: "0.000", 7: "0.00"}, label_cols=(0, 1, 2))
    s.gap()
    s.kv([
        ("sum", "Sum log-prob (response)", f"=SUM(H{a}:H{b})", "0.000", "log π(y|x). Always ≤ 0."),
        ("n", "Trained tokens |y|", f"=SUM(E{a}:E{b})", "0", "prompt tokens are masked out"),
        ("mean", "Mean log-prob", "=«sum»/«n»", "0.000", "what grpo_trainer_v2.py uses"),
        (None, "P(response) = exp(sum)", "=EXP(«sum»)", "0.0000", "probability of the whole answer"),
        (None, "Product of p (check)", f"=PRODUCT(I{a}:I{b})", "0.0000", "same number, via the product"),
        (None, "Perplexity", "=EXP(-«mean»)", "0.00", "≈ 1 means the model was confident"),
    ])
    s.note("What to notice: the ' +' token (p = 0.40) is the uncertain one. It alone costs ln 0.40 = −0.92, "
           "more than all the other tokens together. RL raising that one probability is exactly 'learning the fix'.")
    return s


def s05(wb):
    s = new(wb, "05 Policy Gradient", "Foundations", "Policy gradient (REINFORCE)",
            "We can't differentiate through the tests, so we raise the log-probability of whatever we sampled, weighted by how good it was.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "The reward comes from running tests, a black box with no gradient. The policy-gradient trick avoids "
        "needing one: adjust the log-probability of the answer we sampled, scaled by its reward.",
        "Good sample → push its probability up. Reward 0 → no push at all (the problem sheet 06 fixes).",
        "For a softmax over choices, the gradient of log π(a) with respect to the logits is (one-hot(a) − π). "
        "It raises the chosen logit and lowers every other logit in proportion to its probability.",
        "An LLM does this per token across billions of weights, but the math is identical to this 4-choice toy. "
        "swe-rl-ipr/ipr/train/run.py trains exactly this kind of tabular softmax.",
        "In code: loss = −R × log π(y), then gradient DEScent on the loss = gradient AScent on reward.",
    ])
    s.section("MATH", "math")
    s.math([
        "∇_θ J = E_{y ~ π_θ} [ R(y) · ∇_θ log π_θ(y) ]        (log-derivative trick: ∇π = π · ∇log π)",
        "Softmax policy with logits z:   ∂ log π(a) / ∂ z_k = 1[k = a] − π_k",
        "Exact expected gradient:   ∂J / ∂z_k = π_k · (R_k − J)",
        "Update (gradient ascent):   z_k ← z_k + η · ∂J / ∂z_k",
    ])
    s.legend([("z_k", "logit (score) of choice k"), ("π_k", "softmax probability of choice k"),
              ("η", "learning rate"), ("J", "expected reward = Σ π_k R_k")])
    s.section("EXAMPLE A: one exact gradient step on 4 ways to fix sum_range", "example")
    s.kv([("eta", "Learning rate η", 1.0, "0.00", "try 3.0 to see a bigger jump")])
    data = [("Correct: range(a, b+1)", 1, 0.5), ("Unchanged: range(a, b)", 0, 1.0),
            ("Correct: formula", 1, 0.0), ("New bug: range(a+1, b+1)", 0, 0.2)]
    a, b = s.r + 1, s.r + 4
    tot = b + 1

    def row(i, r):
        if i < 4:
            return [data[i][0], data[i][1], data[i][2], f"=EXP(D{r})/SUMPRODUCT(EXP($D${a}:$D${b}))",
                    f"=C{r}*E{r}", f"=E{r}*(C{r}-$F${tot})", f"=D{r}+«eta»*G{r}",
                    f"=EXP(H{r})/SUMPRODUCT(EXP($H${a}:$H${b}))", f"=C{r}*I{r}"]
        return ["Total", "", "", f"=SUM(E{a}:E{b})", f"=SUM(F{a}:F{b})", f"=SUM(G{a}:G{b})", "",
                f"=SUM(I{a}:I{b})", f"=SUM(J{a}:J{b})"]

    s.table(["Strategy", "Reward R", "Logit z", "π = softmax(z)", "π × R", "Gradient π(R − J)", "New logit",
             "New π", "New π × R"], 5, row,
            fmts={1: "0", 2: "0.00", 3: "0.000", 4: "0.000", 5: "0.000", 6: "0.00", 7: "0.000", 8: "0.000"},
            bold_rows=(4,))
    s.note("Read the Total row: J before is under 'π × R' and J after is under 'New π × R'. One step raises J. Both "
           "correct strategies gain probability and both wrong ones lose it. The gradients sum to 0 because "
           "probability only moves between choices.")
    s.section("EXAMPLE B: a single sampled rollout (what training actually sees)", "example")
    s.kv([("k", "Sampled strategy # (1–4)", 2, "0", "which answer the model happened to write"),
          ("Rk", "Its reward", f"=INDEX($C${a}:$C${b},«k»)", "0", "")])
    s.table(["Strategy", "#", "π", "One-hot of sample", "Gradient R × (one-hot − π)"], 4,
            lambda i, r: [f"=B{a + i}", i + 1, f"=E{a + i}", f"=IF(C{r}=«k»,1,0)", f"=«Rk»*(E{r}-D{r})"],
            fmts={1: "0", 2: "0.000", 3: "0", 4: "0.000"}, label_cols=(0, 1))
    s.note("What to notice: with the default sample #2 (reward 0) every gradient is 0. A failure teaches nothing, "
           "and the model is never told 'don't do that'. Set it to 1 to see a success push its own logit up. "
           "Sheet 06 fixes this with a baseline.")
    return s


def s06(wb):
    s = new(wb, "06 Baseline & Advantage", "Foundations", "Baselines and advantage",
            "Score each answer relative to what was expected (advantage = R − baseline), so failures push down and learning gets far less noisy.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Plain REINFORCE with 0/1 rewards ignores failures, and it still pushes up successes that were already "
        "expected.",
        "Subtract a baseline b (the expected reward): advantage A = R − b. Better than expected → A > 0 (push up). "
        "Worse than expected → A < 0 (push down).",
        "The expected gradient is unchanged (the baseline is unbiased as long as b doesn't depend on the sampled "
        "answer), but the variance drops a lot, so learning is faster and more stable.",
        "PPO learns b with a separate value network (a critic), which is another large model to train and store. "
        "GRPO uses the average reward of the group instead (sheet 07).",
    ])
    s.section("MATH", "math")
    s.math([
        "∇J = E[ (R − b) · ∇ log π(y) ]",
        "Why unbiased:  E[ b · ∇ log π ] = b · Σ_y π(y) ∇ log π(y) = b · ∇ Σ_y π(y) = b · ∇1 = 0",
        "Single-sample estimate for logit k:  g = (R_a − b) · (1[a = k] − π_k)",
        "Var[g] = E[g²] − (E[g])²",
    ])
    s.section("EXAMPLE: same expected gradient, much lower variance", "example")
    s.kv([("k", "Logit we study (k)", 1, "0", "gradient estimate for strategy #k"),
          ("J", "Baseline b = J = Σ π R", "=SUMPRODUCT(«pi»,«R»)", "0.000", "the expected reward")])
    data = [("Correct: range(a, b+1)", 0.4, 1), ("Unchanged bug", 0.3, 0), ("Correct: formula", 0.2, 1),
            ("New bug", 0.1, 0)]
    a, b = s.r + 1, s.r + 4
    s.table(["Sampled strategy a", "#", "π(a)", "Reward R(a)", "g with b = 0", "g with b = J"], 4,
            lambda i, r: [data[i][0], i + 1, data[i][1], data[i][2],
                          f"=(E{r}-0)*(IF(C{r}=«k»,1,0)-INDEX($D${a}:$D${b},«k»))",
                          f"=(E{r}-«J»)*(IF(C{r}=«k»,1,0)-INDEX($D${a}:$D${b},«k»))"],
            fmts={1: "0", 2: "0.00", 3: "0", 4: "0.000", 5: "0.000"}, label_cols=(0, 1))
    s.names["pi"] = f"$D${a}:$D${b}"
    s.names["R"] = f"$E${a}:$E${b}"
    s.gap()
    s.kv([
        ("E0", "E[g], no baseline", f"=SUMPRODUCT(D{a}:D{b},F{a}:F{b})", "0.0000", ""),
        ("EJ", "E[g], baseline J", f"=SUMPRODUCT(D{a}:D{b},G{a}:G{b})", "0.0000", "identical → still unbiased"),
        ("V0", "Var[g], no baseline", f"=SUMPRODUCT(D{a}:D{b},F{a}:F{b},F{a}:F{b})-«E0»^2", "0.0000", ""),
        ("VJ", "Var[g], baseline J", f"=SUMPRODUCT(D{a}:D{b},G{a}:G{b},G{a}:G{b})-«EJ»^2", "0.0000", ""),
        (None, "Variance reduction ×", "=«V0»/«VJ»", "0.0", "how many times less noisy each sample is"),
    ])
    s.note("What to notice: both estimators average to the same gradient (0.16), but the baseline version is about "
           "6× less noisy. Every rollout now carries information, including failures, which get g ≠ 0.")
    return s


def s07(wb):
    s = new(wb, "07 GRPO Advantage", "GRPO mechanics", "GRPO group-relative advantage",
            "Sample G answers to the same prompt. Each answer's advantage is its reward minus the group mean, divided by the group std.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "GRPO (Group Relative Policy Optimization, from DeepSeekMath 2024) uses the group itself as the baseline. "
        "Sample G answers to one prompt, and compare each answer with its siblings.",
        "No critic network: half the memory and much simpler. It works well when reward only arrives at the end.",
        "Dividing by the group std puts every prompt on the same scale, so easy and hard prompts contribute "
        "similar-sized updates.",
        "If all G rewards are equal (all solved or all failed), std = 0 → every advantage is 0 → the prompt is "
        "skipped (grpo_update returns skipped=True).",
        "Rare successes get large positive advantages: a surprising success gets a strong push. v2 clips "
        "advantages at ±5 for stability.",
    ])
    s.sub("Where it lives in the repo")
    s.para("grpo_trainer_v2.py grpo_update(): mean_r, std_r, advantages, advantage_clip;  "
           "grpo_rollouts.py collect_rollouts() prints the same numbers.")
    s.section("MATH", "math")
    s.math([
        "A_i = (R_i − mean(R)) / (std(R) + ε)          population std (divide by G); v2 skips the group when std < 1e-8, swe-rl-ipr adds ε = 1e-6",
        "With binary rewards and k successes out of G, write p = k / G:",
        "     success:  A = (1 − p) / √(p(1 − p)) = √((1 − p) / p)          failure:  A = − √(p / (1 − p))",
        "Σ_i A_i = 0 always: probability is moved from failures to successes.",
    ])
    s.section("EXAMPLE A: one group of 8 rollouts (edit the rewards)", "example")
    s.kv([("clip", "Advantage clip", 5, "0.0", "grpo_trainer_v2.py Config.advantage_clip")])
    rewards = [1, 0, 0, 1, 0, 0, 1, 0]
    a, b = s.r + 1, s.r + 8
    s.table(["Rollout", "Reward R (0/1)", "R − mean", "Advantage A", "Clipped A", "What GRPO does"], 8,
            lambda i, r: [f"Rollout {i + 1}", rewards[i], f"=C{r}-«mean»",
                          f"=IF(«std»<0.00000001,0,D{r}/«std»)", f"=MAX(-«clip»,MIN(«clip»,E{r}))",
                          f'=IF(E{r}>0,"push probability UP",IF(E{r}<0,"push probability DOWN","no change"))'],
            fmts={1: "0", 2: "0.000", 3: "0.000", 4: "0.000"})
    s.gap()
    s.kv([
        ("mean", "Group mean", f"=AVERAGE(C{a}:C{b})", "0.000", "the baseline"),
        ("std", "Group std (population)", f"=STDEVP(C{a}:C{b})", "0.000", "0 → skip this prompt"),
        ("G", "Group size G", f"=COUNT(C{a}:C{b})", "0", ""),
        ("p", "Solve rate p", f"=SUM(C{a}:C{b})/«G»", "0.000", ""),
        (None, "Σ advantages (check = 0)", f"=SUM(E{a}:E{b})", "0.000", ""),
        (None, "Closed form: success A", '=IF(OR(«p»=0,«p»=1),0,SQRT((1-«p»)/«p»))', "0.000", "matches the table"),
        (None, "Closed form: failure A", '=IF(OR(«p»=0,«p»=1),0,-SQRT(«p»/(1-«p»)))', "0.000", ""),
    ])
    s.section("EXAMPLE B: every possible outcome for a group of G = 8", "example")
    f1, f2 = s.table(["Successes k", "p = k / 8", "Success advantage", "Failure advantage", "Learning signal?"], 9,
                     lambda i, r: [i, f"=B{r}/8", f'=IF(OR(C{r}=0,C{r}=1),0,SQRT((1-C{r})/C{r}))',
                                   f'=IF(OR(C{r}=0,C{r}=1),0,-SQRT(C{r}/(1-C{r})))',
                                   f'=IF(OR(B{r}=0,B{r}=8),"none: skipped","yes")'],
                     fmts={0: "0", 1: "0.000", 2: "0.000", 3: "0.000"})
    s.note("What to notice: 1 success in 8 → that success gets A = +2.65 (big push). 7 successes in 8 → the lone "
           "failure gets A = −2.65. The rarer the outcome, the stronger the signal.")
    ch = line_chart("Advantage by number of successes (G = 8)", "successes k", "advantage")
    ch.add_data(Reference(s.ws, min_col=FIRST + 2, max_col=FIRST + 3, min_row=f1 - 1, max_row=f2),
                titles_from_data=True)
    ch.set_categories(Reference(s.ws, min_col=FIRST, min_row=f1, max_row=f2))
    s.ws.add_chart(ch, f"B{s.r + 1}")
    s.r += 18
    return s


def s08(wb):
    s = new(wb, "08 Sweet Spot", "GRPO mechanics", "The difficulty sweet spot",
            "GRPO only learns from groups that mix successes and failures, so tasks must be neither too easy nor too hard for the model.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Learning happens only on mixed groups. A group that all succeeds or all fails has std = 0 and is skipped.",
        "If the model solves a puzzle with probability p, the chance that a group of G is mixed is "
        "1 − p^G − (1 − p)^G.",
        "Mini-SWE-RL's easy puzzles were solved 100% greedily → no signal. The authors wrote medium puzzles tuned "
        "to 30–70% and trained only on the mixed ones.",
        "A larger G widens the useful band because a rare success is more likely to show up. DeepSWE uses G = 16.",
        "Curriculum: as the model improves, puzzles drift toward p ≈ 1 and must be replaced with harder ones.",
    ])
    s.section("MATH", "math")
    s.math([
        "P(all succeed) = p^G        P(all fail) = (1 − p)^G",
        "P(mixed) = 1 − p^G − (1 − p)^G",
        "Expected useful groups per epoch = Σ_puzzles P(mixed_puzzle)",
    ])
    s.section("EXAMPLE A: chance of a useful group vs solve rate", "example")
    s.kv([("G1", "Group size (series 1)", 4, "0", ""), ("G2", "Group size (series 2)", 8, "0", "Mini-SWE-RL v2"),
          ("G3", "Group size (series 3)", 16, "0", "DeepSWE")])
    f1, f2 = s.table(["Solve rate p", '="P(mixed), G="&«G1»', '="P(mixed), G="&«G2»', '="P(mixed), G="&«G3»'], 21,
                     lambda i, r: [i / 20] + [f"=1-B{r}^«{g}»-(1-B{r})^«{g}»" for g in ("G1", "G2", "G3")],
                     fmts={0: "0%", 1: "0.0%", 2: "0.0%", 3: "0.0%"})
    ch = line_chart("P(group has a learning signal)", "model solve rate p", "P(mixed)", width=14, height=9)
    ch.add_data(Reference(s.ws, min_col=FIRST + 1, max_col=FIRST + 3, min_row=f1 - 1, max_row=f2),
                titles_from_data=True)
    ch.set_categories(Reference(s.ws, min_col=FIRST, min_row=f1, max_row=f2))
    s.ws.add_chart(ch, f"G{f1}")
    s.section("EXAMPLE B: the repo's 16 training puzzles (measured solve counts)", "example")
    s.note("Source: comments next to TRAIN_PUZZLES in grpo_trainer_v2.py (hard set measured with 8 attempts, "
           "medium set with 4, both at temperature 0.8).")
    data = [("hard_scope_bug", "hard", 2, 8), ("hard_spiral_order", "hard", 4, 8), ("hard_lru_cache", "hard", 5, 8),
            ("hard_merge_intervals", "hard", 5, 8), ("hard_graph_cycle", "hard", 2, 8),
            ("hard_eval_rpn", "hard", 5, 8), ("hard_knapsack", "hard", 6, 8), ("hard_off_by_one", "hard", 7, 8),
            ("med_recursive_power", "medium", 3, 4), ("med_case_dedup", "medium", 2, 4),
            ("med_class_shared_state", "medium", 1, 4), ("med_sliding_window_avg", "medium", 2, 4),
            ("med_flatten_depth", "medium", 2, 4), ("med_running_total", "medium", 1, 4),
            ("med_nested_depth", "medium", 1, 4), ("med_zip_truncation", "medium", 1, 4)]
    a, b = s.r + 1, s.r + 16
    s.table(["Puzzle", "Set", "Solved", "Attempts", "p", "P(mixed) at G2", "Category"], 16,
            lambda i, r: list(data[i]) + [f"=D{r}/E{r}", f"=1-F{r}^«G2»-(1-F{r})^«G2»",
                                          f'=IF(F{r}=0,"all fail",IF(F{r}=1,"all solved","mixed"))'],
            fmts={2: "0", 3: "0", 4: "0%", 5: "0.0%"}, label_cols=(0, 1))
    s.gap()
    s.kv([(None, "Expected useful groups per epoch", f"=SUM(G{a}:G{b})", "0.0",
           "out of 16 puzzles, at the G in series 2"),
          (None, "Epoch solve rate (history.json, epoch 1)", 0.546875, "0.0%",
           "from checkpoints/grpo_v2/history.json: 70 / 128 rollouts solved")])
    s.note("What to notice: hard_off_by_one (7/8) is mixed only about 66% of the time at G = 8, so a third of its "
           "compute is wasted. Puzzles near 50% almost always teach something.")
    return s


def s09(wb):
    s = new(wb, "09 GRPO Loss", "GRPO mechanics", "The GRPO loss and one update step",
            "Loss = − average over rollouts of (advantage × mean token log-prob). Minimising it lifts good answers and sinks bad ones.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Putting sheets 04 + 07 together: for each rollout, multiply its advantage by its mean token log-prob, "
        "average over the group, and negate. Minimising that raises the log-prob of A > 0 answers and lowers A < 0 ones.",
        "The loss value is NOT a progress meter. It hovers near 0 and can be negative (v2 history.json: avg_loss is −0.010 in epoch 1, reaches −0.080 in epoch 9 and ends at −0.076). "
        "Track reward / solve rate instead.",
        "Memory trick in v2: call backward() after each rollout (gradients accumulate and each graph is freed), "
        "divide by the count, clip the gradient norm to 1.0, and take ONE AdamW step per puzzle.",
        "Learning rate is tiny (5e-6 in v2, 1e-6 in v1): we nudge a pretrained model rather than retrain it.",
    ])
    s.sub("Where it lives in the repo")
    s.para("grpo_trainer_v2.py grpo_update(): pg_loss = -adv * policy_lp.mean(); rollout_loss.backward(); "
           "p.grad.div_(n_valid); clip_grad_norm_(…, 1.0); optimizer.step().")
    s.section("MATH", "math")
    s.math([
        "L(θ) = − (1/G) Σ_i A_i · (1/|y_i|) Σ_t log π_θ(y_i,t | x, y_i,<t)   [ + β · KL, sheet 10 ]",
        "Change in an answer's probability after an update = exp( |y| · (new mean log-prob − old mean log-prob) )",
    ])
    s.section("EXAMPLE: 4 rollouts, before and after one (illustrative) update", "example")
    s.note("Mean log-probs 'after' are illustrative: a real step moves each one in the direction of its advantage.")
    data = [("Rollout 1", 1, 42, -0.35, -0.30), ("Rollout 2", 0, 55, -0.60, -0.66), ("Rollout 3", 0, 38, -0.42, -0.47),
            ("Rollout 4", 1, 60, -0.80, -0.70)]
    a, b = s.r + 1, s.r + 4
    s.table(["Rollout", "Reward", "Advantage A", "Tokens |y|", "Mean log-prob (before)", "Loss term −A × mean lp (before)",
             "Mean log-prob (after)", "Loss term (after)", "P(answer) change ×"], 4,
            lambda i, r: [data[i][0], data[i][1], f"=IF(«std»=0,0,(C{r}-«mean»)/«std»)", data[i][2], data[i][3],
                          f"=-D{r}*F{r}", data[i][4], f"=-D{r}*H{r}", f"=EXP(E{r}*(H{r}-F{r}))"],
            fmts={1: "0", 2: "0.000", 3: "0", 4: "0.00", 5: "0.000", 6: "0.00", 7: "0.000", 8: "0.00"},
            header_height=44)
    s.gap()
    s.kv([
        ("mean", "Reward mean", f"=AVERAGE(C{a}:C{b})", "0.000", ""),
        ("std", "Reward std", f"=STDEVP(C{a}:C{b})", "0.000", ""),
        (None, "GRPO loss (before)", f"=AVERAGE(G{a}:G{b})", "0.0000", "average of the loss terms"),
        (None, "GRPO loss (after)", f"=AVERAGE(I{a}:I{b})", "0.0000", "lower → the step helped"),
    ])
    s.note("What to notice: rollout 4 (a success with low log-prob) becomes about 400× more likely, and the "
           "failures become about 30× and 7× less likely. Small per-token changes compound over 40–60 tokens.")
    return s


def s10(wb):
    s = new(wb, "10 KL Penalty", "Production recipe", "KL penalty and the reference model",
            "A KL term measures how far the trained model has drifted from a frozen copy of the original. It acts as a leash against forgetting.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Pushing hard on a few puzzles can make the model drift: weird text, lost skills. KL divergence measures how "
        "different the new model's distribution is from a frozen reference copy (the starting weights).",
        "Adding β · KL to the loss penalises drift. Mini-SWE-RL v1 uses β = 0.01. v2 has β = 0.05 but ran with "
        "--no-ref (KL off) to save memory. The README reports 7 puzzles gained and 5 regressed, and blames the regressions on forgetting (its gains table actually lists 8, see sheet 16).",
        "DeepSWE / DAPO-style GRPO++ drop the KL term and rely on clipping (sheet 11) for stability.",
        "Summing over the whole vocabulary at every token is expensive, so KL is estimated from the sampled tokens. "
        "[From the wider RL literature, not the repo:] k1 = log π − log π_ref is unbiased but noisy and can be negative. k3 = (r − 1) − log r with "
        "r = π_ref / π is unbiased, always ≥ 0, and lower variance (the one used in the GRPO paper).",
        "Repo note: v1 sums and v2 averages π_ref · (log π_ref − log π) over the sampled tokens. That is a non-standard "
        "quantity, not a KL estimate, and it shrinks when the reference probability is small.",
    ])
    s.section("MATH", "math")
    s.math([
        "KL(π ‖ π_ref) = Σ_y π(y) · ln( π(y) / π_ref(y) )   ≥ 0, and 0 only when the distributions are equal",
        "Per sampled token:  k1 = ln π − ln π_ref          k3 = (r − 1) − ln r,  with r = π_ref / π",
        "Loss = GRPO loss + β · mean_t(k)",
    ])
    s.section("EXAMPLE A: exact KL over a 5-token vocabulary", "example")
    data = [("return", 0.50, 0.70), ("if", 0.25, 0.15), ("for", 0.15, 0.08), ("while", 0.07, 0.05), ("pass", 0.03, 0.02)]
    a, b = s.r + 1, s.r + 5
    s.table(["Token", "π_ref (original)", "π (trained)", "π · ln(π/π_ref)", "π_ref · ln(π_ref/π)"], 5,
            lambda i, r: list(data[i]) + [f"=D{r}*LN(D{r}/C{r})", f"=C{r}*LN(C{r}/D{r})"],
            fmts={1: "0.00", 2: "0.00", 3: "0.0000", 4: "0.0000"})
    s.gap()
    s.kv([(None, "KL(π ‖ π_ref)", f"=SUM(E{a}:E{b})", "0.0000", "the one used in RLHF / GRPO"),
          (None, "KL(π_ref ‖ π)", f"=SUM(F{a}:F{b})", "0.0000", "KL is not symmetric")])
    s.section("EXAMPLE B: estimating KL from 6 sampled tokens", "example")
    s.kv([("beta", "β (KL coefficient)", 0.05, "0.00", "grpo_trainer_v2.py Config.kl_coef")])
    data = [("return", 0.70, 0.50), (" a", 0.90, 0.88), (" +", 0.65, 0.40), (" b", 0.95, 0.94), ("\\n", 0.80, 0.85),
            ("<eos>", 0.92, 0.90)]
    a, b = s.r + 1, s.r + 6
    s.table(["Sampled token", "π (trained)", "π_ref", "k1 = ln π − ln π_ref", "r = π_ref / π", "k3 = r − 1 − ln r",
             "Repo formula π_ref(ln π_ref − ln π)"], 6,
            lambda i, r: list(data[i]) + [f"=LN(C{r})-LN(D{r})", f"=D{r}/C{r}", f"=F{r}-1-LN(F{r})",
                                          f"=D{r}*(LN(D{r})-LN(C{r}))"],
            fmts={1: "0.00", 2: "0.00", 3: "0.0000", 4: "0.000", 5: "0.0000", 6: "0.0000"}, header_height=44)
    s.gap()
    s.kv([(None, "Mean k1", f"=AVERAGE(E{a}:E{b})", "0.0000", "can go negative on single tokens"),
          ("k3", "Mean k3", f"=AVERAGE(G{a}:G{b})", "0.0000", "always ≥ 0"),
          (None, "Mean repo formula", f"=AVERAGE(H{a}:H{b})", "0.0000", "negative here, so it would REWARD drift"),
          (None, "KL penalty added to loss = β · mean k3", "=«beta»*«k3»", "0.00000", "")])
    return s


def s11(wb):
    s = new(wb, "11 PPO Clipping", "Production recipe", "Importance ratio and PPO / clip-higher",
            "Cap how far one batch can move each token's probability. Clip-higher gives more room to raise rare good tokens.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Trainers usually take several gradient steps on one batch of rollouts (swe-rl-ipr config.yaml: train_batch_size 64 with policy_mini_batch_size 16, i.e. 4 mini-batch steps per batch). After the first step, the model "
        "that generated the data (π_old) is no longer the model being trained (π_θ).",
        "The importance ratio r = π_θ / π_old (per token) corrects for that mismatch.",
        "Clipping keeps r inside [1 − ε_low, 1 + ε_high] in the objective. Once a token's probability has moved "
        "far enough in the right direction, its gradient switches off. This is the trust region that keeps PPO / "
        "GRPO stable.",
        "Clip-higher (DAPO, used in DeepSWE's GRPO++): ε_high = 0.28 > ε_low = 0.20. Rare good tokens get more room "
        "to grow, which keeps exploration alive and prevents entropy collapse.",
        "Mini-SWE-RL takes one step per batch with no ratio (implicitly r = 1). That is only valid if the data "
        "is on-policy, which v2's data isn't (sheet 14).",
    ])
    s.sub("Where it lives in the repo")
    s.para("swe-rl-ipr/ipr/train/config.yaml: eps_clip_low 0.20, eps_clip_high 0.28, use_kl_loss false.")
    s.section("MATH", "math")
    s.math([
        "r_t = π_θ(y_t | …) / π_old(y_t | …) = exp( log π_θ − log π_old )",
        "objective_t = min( r_t · A ,  clip(r_t, 1 − ε_low, 1 + ε_high) · A )          loss = − mean_t objective_t",
        "Gradient flows only when the min picks the un-clipped term.",
    ])
    s.section("EXAMPLE: objective vs ratio for a good (A > 0) and a bad (A < 0) answer", "example")
    s.kv([("el", "ε_low", 0.20, "0.00", ""), ("eh", "ε_high", 0.28, "0.00", "clip-higher"),
          ("Ap", "Advantage of a good answer", 1.0, "0.00", ""), ("An", "Advantage of a bad answer", -1.0, "0.00", "")])
    ratios = [0.5, 0.7, 0.8, 0.9, 1.0, 1.1, 1.2, 1.28, 1.4, 1.6]
    f1, f2 = s.table(["Ratio r", "clip(r)", "A>0: r·A", "A>0: clip·A", "A>0: objective", "A>0: gradient?",
                      "A<0: r·A", "A<0: clip·A", "A<0: objective", "A<0: gradient?"], len(ratios),
                     lambda i, r: [ratios[i], f"=MIN(MAX(B{r},1-«el»),1+«eh»)", f"=B{r}*«Ap»", f"=C{r}*«Ap»",
                                   f"=MIN(D{r},E{r})", f'=IF(D{r}<=E{r}+0.000000001,"yes","no (clipped)")',
                                   f"=B{r}*«An»", f"=C{r}*«An»", f"=MIN(H{r},I{r})",
                                   f'=IF(H{r}<=I{r}+0.000000001,"yes","no (clipped)")'],
                     fmts={0: "0.00", 1: "0.00", 2: "0.00", 3: "0.00", 4: "0.00", 6: "0.00", 7: "0.00", 8: "0.00"},
                     header_height=30)
    s.note("What to notice: a good answer stops getting pushed once r > 1.28 (already boosted enough). A bad answer "
           "stops getting pushed down once r < 0.80. When a bad answer's ratio went UP, the gradient stays on, so "
           "mistakes are always corrected (the pessimistic min).")
    ch = line_chart("Clipped objective vs ratio", "ratio r", "objective")
    for c in (FIRST + 4, FIRST + 8):
        ch.add_data(Reference(s.ws, min_col=c, min_row=f1 - 1, max_row=f2), titles_from_data=True)
    ch.set_categories(Reference(s.ws, min_col=FIRST, min_row=f1, max_row=f2))
    s.ws.add_chart(ch, f"B{s.r + 1}")
    s.r += 18
    return s


def s12(wb):
    s = new(wb, "12 Leave-One-Out", "Production recipe", "Leave-one-out (RLOO) baseline",
            "Each answer's baseline is the average of the OTHER answers, so it doesn't include its own reward.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "The group mean includes the answer itself, so each baseline is slightly contaminated by its own reward.",
        "RLOO: baseline_i = average reward of the other G − 1 answers. The baseline is then independent of the "
        "answer it scores (strictly unbiased).",
        "Identity: R_i − b_i(LOO) = G / (G − 1) · (R_i − mean). With std-normalisation, LOO only rescales every "
        "advantage by the same factor (8/7 ≈ 1.14 for G = 8). Directions and ranking are unchanged.",
        "The practical difference is small for G = 8 and larger for tiny groups (G = 2: factor 2).",
    ])
    s.sub("Where it lives in the repo")
    s.para("swe-rl-ipr/ipr/train/grpo.py leave_one_out_baseline(), grpo_advantages();  ipr/train/run.py _grpo_advantages().")
    s.section("MATH", "math")
    s.math([
        "b_i(LOO) = ( Σ_j R_j − R_i ) / (G − 1)",
        "A_i(LOO) = R_i − b_i(LOO) = G/(G−1) · (R_i − mean(R))",
    ])
    s.section("EXAMPLE: group mean vs leave-one-out on 8 rollouts", "example")
    rewards = [1, 1, 0, 0, 0, 1, 0, 0]
    a, b = s.r + 1, s.r + 8
    s.table(["Rollout", "Reward R", "Mean baseline", "A = R − mean", "LOO baseline", "A(LOO) = R − b", "A(LOO) / A",
             "A / std", "A(LOO) / std"], 8,
            lambda i, r: [f"Rollout {i + 1}", rewards[i], "=«mean»", f"=C{r}-D{r}", f"=(«sum»-C{r})/(«G»-1)",
                          f"=C{r}-F{r}", f'=IF(ABS(E{r})<0.000000001,"–",G{r}/E{r})',
                          f"=IF(«std»=0,0,E{r}/«std»)", f"=IF(«std»=0,0,G{r}/«std»)"],
            fmts={1: "0", 2: "0.000", 3: "0.000", 4: "0.000", 5: "0.000", 6: "0.000", 7: "0.000", 8: "0.000"})
    s.gap()
    s.kv([("G", "G", f"=COUNT(C{a}:C{b})", "0", ""), ("sum", "Σ R", f"=SUM(C{a}:C{b})", "0", ""),
          ("mean", "mean R", "=«sum»/«G»", "0.000", ""), ("std", "std R", f"=STDEVP(C{a}:C{b})", "0.000", ""),
          (None, "G / (G − 1)", "=«G»/(«G»-1)", "0.000", "the constant ratio in column H")])
    s.note("Repo detail: ipr/train/grpo.py divides the LOO advantage by the std of the rewards (column J = column I × G/(G−1)). "
           "ipr/train/run.py divides by the std of the LOO advantages themselves, which cancels the factor and gives exactly column I.")
    return s


def s13(wb):
    s = new(wb, "13 Length Norm & Filter", "Production recipe", "Length normalisation and compact filtering",
            "How token losses are averaged decides whether long answers are favoured. Rollouts cut off by a limit should teach nothing.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Three ways to combine token losses: (a) SUM over tokens (v1): long answers get bigger updates. "
        "(b) per-answer MEAN (original GRPO, v2): every answer counts once, so each token of a short answer "
        "weighs more than a token of a long one. (c) divide by a CONSTANT such as the max length (Dr. GRPO / "
        "DeepSWE; ipr/train/grpo.py says it divides by the max context length): every token weighs the same. Dr. GRPO is a paper from the wider literature, not cited in the repo.",
        "Length bias of (b): with a negative advantage, a long wrong answer is penalised less per token than a "
        "short wrong one, so wrong answers drift longer over training.",
        "Compact filtering (DeepSWE): rollouts that hit the max length, the time limit or the step limit get "
        "advantage 0. Otherwise a good-but-unfinished attempt is punished like a failure, which teaches the model "
        "to give up early.",
    ])
    s.sub("Where it lives in the repo")
    s.para("grpo_trainer.py: policy_log_probs.sum() (a);  grpo_trainer_v2.py: policy_lp.mean() (b);  "
           "swe-rl-ipr/ipr/train/grpo.py: length_normalize, compact_filtering, score_rollout(masked=…).")
    s.section("MATH", "math")
    s.math([
        "Per-token weight:   SUM → A_i        MEAN → A_i / |y_i|        CONSTANT → A_i / L_max",
        "Compact filtering:  A_i := 0 if rollout i was truncated / timed out / hit the step cap",
    ])
    s.section("EXAMPLE: four rollouts under each scheme", "example")
    s.kv([("Lmax", "Max length L_max (tokens)", 512, "0", ""),
          ("filt", "Compact filtering on? (1 = yes, 0 = no)", 1, "0", "")])
    data = [("Short correct", 20, 1.2, 0), ("Long correct", 200, 1.2, 0), ("Long wrong", 300, -0.8, 0),
            ("Cut off at L_max", 512, -0.8, 1)]
    s.table(["Rollout", "Tokens |y|", "Advantage A", "Truncated? (1/0)", "Effective A", "Per-token weight: SUM",
             "Per-token weight: MEAN", "Per-token weight: CONST", "Total push: MEAN", "Total push: CONST"], 4,
            lambda i, r: list(data[i]) + [f"=IF(AND(«filt»=1,E{r}=1),0,D{r})", f"=F{r}", f"=F{r}/C{r}",
                                          f"=F{r}/«Lmax»", f"=H{r}*C{r}", f"=I{r}*C{r}"],
            fmts={1: "0", 2: "0.00", 3: "0", 4: "0.00", 5: "0.000", 6: "0.0000", 7: "0.0000", 8: "0.00", 9: "0.00"},
            header_height=44)
    s.note("What to notice: under MEAN, each token of 'Short correct' gets 10× the weight of a 'Long correct' token, "
           "and 'Long wrong' is penalised only −0.0027 per token. Under CONST all tokens are treated alike. "
           "Turn filtering off (0) to see the cut-off rollout get punished as a failure.")
    return s


def s14(wb):
    s = new(wb, "14 On vs Off Policy", "Production recipe", "On-policy vs off-policy data",
            "Policy gradients assume the data came from the model being trained. Mini-SWE-RL v2 breaks that assumption.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "On-policy: the rollouts used for a gradient step were generated by the current model. The formulas on "
        "sheets 05–09 assume this.",
        "Mini-SWE-RL v2 is a hybrid. Ollama (frozen original weights, likely quantised) generates every rollout, "
        "and a separate HuggingFace copy is trained. After the first update the trainee is not the generator, the "
        "gap grows every step, and there is no importance-ratio correction.",
        "So v2 behaves like reward-weighted fine-tuning on the base model's samples. It can never learn from its "
        "own new behaviour.",
        "Template mismatch too: Ollama's /api/generate applies the model's chat template by default (Ollama behaviour, not repo code), but the trainer "
        "scores raw prompt + response text. The 'before' eval uses Ollama and the 'after' eval uses HuggingFace "
        "on the raw prompt.",
        "Fixes: generate with the model being trained (as v1 does), sync weights to the inference server (vLLM "
        "weight sync in production), or correct with importance ratios + clipping (sheet 11). Keep templates "
        "identical in generation, training and evaluation.",
    ])
    s.section("MATH", "math")
    s.math([
        "E_{y~π_θ}[ A · ∇ log π_θ(y) ] = E_{y~μ}[ (π_θ(y) / μ(y)) · A · ∇ log π_θ(y) ]      μ = behaviour (generator) policy",
        "Sequence ratio = exp( Σ_t (log π_θ(y_t) − log μ(y_t)) )  → explodes or vanishes for long answers,",
        "which is why practical methods use per-token ratios with clipping.",
    ])
    s.section("EXAMPLE A: 5 rollouts generated by μ, scored for the current π_θ", "example")
    s.kv([("el", "ε_low", 0.20, "0.00", ""), ("eh", "ε_high", 0.28, "0.00", "")])
    data = [("Rollout 1", -12.0, -11.5, 1.3), ("Rollout 2", -15.0, -16.2, -0.8), ("Rollout 3", -9.5, -9.4, 1.3),
            ("Rollout 4", -20.0, -18.8, -0.8), ("Rollout 5", -11.0, -13.0, -0.8)]
    s.table(["Rollout", "log μ(y) (generator)", "log π_θ(y) (current)", "Advantage A", "log ratio", "ratio π_θ/μ",
             "Uncorrected weight A", "Corrected ratio × A", "Clipped ratio × A"], 5,
            lambda i, r: list(data[i]) + [f"=D{r}-C{r}", f"=EXP(F{r})", f"=E{r}", f"=G{r}*E{r}",
                                          f"=MIN(MAX(G{r},1-«el»),1+«eh»)*E{r}"],
            fmts={1: "0.0", 2: "0.0", 3: "0.00", 4: "0.00", 5: "0.000", 6: "0.00", 7: "0.00", 8: "0.00"},
            header_height=30)
    s.note("What to notice: rollout 4 is a failure the current model already likes 3.3× more than the generator did. "
           "Uncorrected training under-weights it (−0.8 instead of about −2.7). Rollout 5, which the model has moved "
           "away from, is over-weighted.")
    s.section("EXAMPLE B: how fast a sequence ratio drifts", "example")
    s.kv([("d", "Per-token log-prob drift per update", 0.002, "0.000", "tiny change per token"),
          ("len", "Response length (tokens)", 150, "0", "")])
    s.table(["Updates since generation", "Sequence log ratio", "Sequence ratio"], 11,
            lambda i, r: [i, f"=B{r}*«d»*«len»", f"=EXP(C{r})"], fmts={0: "0", 1: "0.00", 2: "0.00"})
    s.note("After 10 updates a 0.002-per-token drift makes whole-answer ratios about 20×, so frozen-generator "
           "data goes stale fast.")
    return s


def s15(wb):
    s = new(wb, "15 Multi-Turn Agents", "Production recipe", "Multi-turn agents (the real SWE setting)",
            "Real SWE agents loop through think → tool → observe for many turns and get one reward at the end. Only the model's own tokens are trained.",
            widths={"D": 28})
    s.section("CONCEPT", "concept")
    s.bullets([
        "DeepSWE, SWE-agent and OpenHands work in a loop: call a tool (search, open file, edit, run tests), read "
        "the result, and repeat for many steps before submitting a patch (swe-rl-ipr's config caps it at 100, noting a typical horizon of 50–100).",
        "The whole trajectory gets ONE reward at the end (tests pass or not). GRPO still works: a group is "
        "G full trajectories on the same issue.",
        "Only tokens the model wrote are trained. Tool outputs (file contents, test logs) are masked, because the "
        "model shouldn't learn to predict the environment.",
        "New problems at this scale: very long contexts (32k tokens), slow rollouts (Docker + test suites) that "
        "need async generation, credit assignment across many steps, and timeouts (handled by compact filtering).",
    ])
    s.sub("Where it lives in the repo")
    s.para("swe-rl-ipr/ipr/harness/agent.py (ReAct loop, OpenAIChatPolicy for vLLM);  harness/tools.py "
           "(bash, search, open, edit, run_sanctioned_tests, submit);  config.yaml max_agent_steps 100, max_context_len 32768.")
    s.section("MATH", "math")
    s.math([
        "Trajectory τ = (x, a_1, o_1, a_2, o_2, …, a_T)          a = model actions, o = tool observations",
        "log π(τ) = Σ_turns Σ_tokens in a_t log π(token)          (observation tokens have mask m = 0)",
        "L = − A(τ) · Σ_t m_t · log π(token_t) / normaliser",
    ])
    s.section("EXAMPLE: token budget of one trajectory (illustrative token counts)", "example")
    s.kv([("ctx", "Max context (tokens)", 32768, "#,##0", "config.yaml max_context_len")])
    data = [("System", "Tool specs + instructions", 400), ("User", "GitHub issue text", 350),
            ("Assistant", "search('parse_version')", 60), ("Tool", "12 matching lines", 500),
            ("Assistant", "open('utils/version.py')", 40), ("Tool", "file contents", 900),
            ("Assistant", "edit(old → new)", 180), ("Tool", "edit applied", 30),
            ("Assistant", "run_sanctioned_tests()", 30), ("Tool", "test log", 1200), ("Assistant", "submit()", 20)]
    a, b = s.r + 1, s.r + len(data)
    s.table(["Turn", "Role", "What happens", "Tokens", "Trained? (mask)", "Trained tokens", "Context used so far"],
            len(data),
            lambda i, r: [i + 1, data[i][0], data[i][1], data[i][2], f'=IF(C{r}="Assistant",1,0)', f"=E{r}*F{r}",
                          f"=SUM($E${a}:E{r})"],
            fmts={0: "0", 3: "#,##0", 4: "0", 5: "#,##0", 6: "#,##0"}, label_cols=(0, 1, 2))
    s.gap()
    s.kv([("tot", "Total tokens", f"=SUM(E{a}:E{b})", "#,##0", ""),
          ("tr", "Trained tokens", f"=SUM(G{a}:G{b})", "#,##0", ""),
          (None, "Share of tokens trained", "=«tr»/«tot»", "0.0%", "most of the context is environment output"),
          (None, "Context budget used", "=«tot»/«ctx»", "0.0%", "a 50-step run can hit the cap → compact filtering")])
    return s


def s16(wb):
    s = new(wb, "16 Evaluation & Noise", "Evaluation & reward integrity", "Evaluation and statistical noise",
            "Evaluate the same way before and after, on held-out tasks, and check whether the change is bigger than noise.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Use greedy decoding (T = 0) on a fixed task set, before and after training, with the SAME setup (engine, "
        "prompt template, code extraction). Otherwise you measure the setup change, not learning.",
        "Hold out tasks the model never trained on. In Mini-SWE-RL v2, 16 of the 30 evaluation puzzles are "
        "training puzzles.",
        "Small evaluation sets are noisy: with 30 puzzles, one puzzle = 3.3 points.",
        "Paired comparison: look at the puzzles that flipped. If training did nothing, each flip is equally likely "
        "to be a gain or a loss (sign test / McNemar).",
        "Mini-SWE-RL: 7 gained, 5 lost (net +2) → p ≈ 0.77, indistinguishable from noise. swe-rl-ipr 7B MBPP: "
        "plus@1 rising 0.633 → 0.667 under IPR (flat at 0.633 under the raw reward) is a small change. The repo doesn't state the eval-split size, but every reported value is a multiple of 1/30, which suggests about 1 problem out of 30.",
        "Better practice: more evaluation tasks, several seeds, and pass@k with sampling.",
    ])
    s.section("MATH", "math")
    s.math([
        "Standard error SE = √( p (1 − p) / n )          95% interval ≈ p ± 1.96 · SE",
        "Sign test: X ~ Binomial(gains + losses, 0.5);   two-sided p-value = min(1, 2 · P(X ≤ min(gains, losses)))",
        "Tasks needed to detect a lift Δ (rough, 80% power):  n ≈ (1.96 + 0.84)² · 2p(1 − p) / Δ²",
        "pass@k (unbiased, from n samples with c correct) = 1 − C(n − c, k) / C(n, k)",
    ])
    s.section("EXAMPLE A: is 20/30 → 22/30 real?", "example")
    s.kv([("n", "Evaluation puzzles", 30, "0", "README: 30 puzzles"), ("b0", "Solved before", 20, "0", ""),
          ("b1", "Solved after", 22, "0", ""),
          ("p0", "Rate before", "=«b0»/«n»", "0.0%", ""), ("p1", "Rate after", "=«b1»/«n»", "0.0%", ""),
          ("se0", "SE before", "=SQRT(«p0»*(1-«p0»)/«n»)", "0.0%", ""),
          (None, "95% interval before", '=TEXT(«p0»-1.96*«se0»,"0%")&" – "&TEXT(«p0»+1.96*«se0»,"0%")', None, ""),
          ("se1", "SE after", "=SQRT(«p1»*(1-«p1»)/«n»)", "0.0%", ""),
          (None, "95% interval after", '=TEXT(«p1»-1.96*«se1»,"0%")&" – "&TEXT(«p1»+1.96*«se1»,"0%")', None,
           "the intervals overlap almost completely")])
    s.section("EXAMPLE B: paired sign test on the flipped puzzles", "example")
    s.note("Caution: the README is internally inconsistent. The heading says '7 newly solved' but its table lists 8 "
           "(4 hard + 4 medium), and 4 hard gains − 3 hard losses would not leave the hard score unchanged at 73.3%. "
           "7 vs 5 matches the stated totals (20 → 22). Try 8 vs 5: p ≈ 0.58, still not significant.")
    s.kv([("g", "Puzzles gained (FAIL → PASS)", 7, "0", "README headline"), ("l", "Puzzles lost (PASS → FAIL)", 5, "0", "README"),
          ("pv", "Two-sided p-value", "=MIN(1,2*BINOMDIST(MIN(«g»,«l»),«g»+«l»,0.5,TRUE))", "0.000", ""),
          (None, "Verdict", '=IF(«pv»<0.05,"significant","not significant: could be noise")', None, "")])
    s.section("EXAMPLE C: how many tasks would prove a +6.7-point lift?", "example")
    s.kv([("bp", "Baseline solve rate", 0.667, "0.0%", ""), ("dl", "Lift to detect Δ", 0.067, "0.0%", ""),
          (None, "Tasks needed (approx.)", "=ROUNDUP((1.96+0.84)^2*2*«bp»*(1-«bp»)/«dl»^2,0)", "#,##0",
           "roughly 25× the 30 puzzles used")])
    s.section("EXAMPLE D: pass@k from n samples", "example")
    s.kv([("ns", "Samples per task n", 8, "0", ""), ("c", "Correct samples c", 3, "0", "")])
    s.table(["k", "pass@k"], 8,
            lambda i, r: [i + 1, f"=IF(«ns»-«c»<B{r},1,1-COMBIN(«ns»-«c»,B{r})/COMBIN(«ns»,B{r}))"],
            fmts={0: "0", 1: "0.0%"})
    return s


def s17(wb):
    s = new(wb, "17 Reward Hacking", "Evaluation & reward integrity", "Reward hacking (Goodhart's law)",
            "RL optimises exactly the reward you wrote. If a cheat earns +1, the model will learn the cheat.",
            widths={"C": 30})
    s.section("CONCEPT", "concept")
    s.bullets([
        "Goodhart's law: when a measure becomes a target, it stops being a good measure. The proxy here is 'visible "
        "tests pass'; the true goal is 'the code is correct'.",
        "Typical code hacks: hardcode the expected outputs; a lookup table keyed on the test inputs; special-case "
        "the test; delete or edit the tests; swallow every exception; exit early with success.",
        "Why RL finds them: a hack earns the same +1 as a real fix and is often shorter and easier to produce, so "
        "the gradient cannot tell them apart (see sheet 20).",
        "Reward-hacking gap = visible pass rate − held-out pass rate. A gap that grows during training means the "
        "model is hacking.",
        "Defences: held-out tests, perturbed tests (IPR, sheet 18), tamper detection, sandboxing, LLM "
        "judges / monitors.",
    ])
    s.section("MATH", "math")
    s.math([
        "Raw reward  R_raw = 1[visible tests pass]          True goal  R* = 1[correct on all inputs, tests untouched]",
        "Reward-hacking gap = P(visible pass) − P(held-out pass)",
    ])
    s.section("EXAMPLE A: four candidate patches for to_roman (from swe-rl-ipr/ipr/localgym)", "example")
    data = [("Real fix", "Correct algorithm", 1, 1, 0), ("Lookup cheat", "Table of the 4 visible answers", 1, 0, 0),
            ("Tamper cheat", "Real code, but deletes a test", 1, 1, 1), ("Buggy", "Unchanged (no IV / IX)", 0, 0, 0)]
    a, b = s.r + 1, s.r + 4
    s.table(["Candidate", "What it does", "Visible tests pass?", "Held-out tests pass?", "Touched test files?",
             "Raw reward", "Honest & correct", "Verdict"], 4,
            lambda i, r: list(data[i]) + [f"=D{r}", f"=IF(AND(E{r}=1,F{r}=0),1,0)",
                                          f'=IF(AND(G{r}=1,H{r}=0),"HACK: paid but not honest/correct","")'],
            fmts={2: "0", 3: "0", 4: "0", 5: "0", 6: "0"}, label_cols=(0, 1), row_height=28, header_height=30)
    s.gap()
    s.kv([(None, "Visible pass rate", f"=AVERAGE(D{a}:D{b})", "0%", ""),
          (None, "Held-out pass rate", f"=AVERAGE(E{a}:E{b})", "0%", ""),
          (None, "Gap", f"=AVERAGE(D{a}:D{b})-AVERAGE(E{a}:E{b})", "0%", "")])
    s.section("EXAMPLE B: the gap during real training (Qwen2.5-Coder-7B, GRPO on MBPP, eval on MBPP+)", "example")
    s.note("Source: swe-rl-ipr/paper/FACTS.md §3d. base@1 = passes visible MBPP tests; plus@1 = passes the large "
           "held-out MBPP+ suite. The 30-step matched numbers also appear in paper/tables/rl_table.tex. "
           "The IPR arm's starting values aren't given. The paper says its gap 'stays flat at 0.133', which doesn't "
           "match the raw arm's 0.067 starting gap, so the two arms' first evaluations may differ.")
    data = [("Raw reward, 100-step run, start", 0.78, 0.68), ("Raw reward, 100-step run, end", 0.82, 0.68),
            ("Raw reward, 30-step matched run, start", 0.70, 0.633),
            ("Raw reward, 30-step matched run, end", 0.833, 0.633), ("IPR reward, 30-step matched run, end", 0.800, 0.667)]
    s.table(["Run", "base@1 (visible)", "plus@1 (held-out)", "Gap"], 5,
            lambda i, r: list(data[i]) + [f"=C{r}-D{r}"], fmts={1: "0.000", 2: "0.000", 3: "0.000"})
    s.note("What to notice: under the raw reward the visible score climbs while true correctness stays flat, so the "
           "model is learning to pass these particular tests. IPR is the only arm whose held-out score rises. "
           "(The differences are about 1 problem, so see sheet 16.)")
    return s


def s18(wb):
    s = new(wb, "18 IPR Reward", "Evaluation & reward integrity", "Isomorphic Perturbation Reward (IPR)",
            "Re-grade every patch on K fresh versions of the tests, with expected answers from a trusted reference. Real fixes still pass; memorised answers fail.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Step 1, tamper check: if the patch's diff touches any test or config file → reward 0.",
        "Step 2, perturb: for k = 1…K, sample new inputs (seeded, reproducible) and compute the expected outputs "
        "with a trusted gold solution that the agent never sees.",
        "Step 3, admission gate: drop inputs where the gold solution itself errors, so a correct fix is never "
        "punished (zero false-negative noise by construction).",
        "Step 4, AND gate: reward 1 only if every one of the K suites passes. The visible suite runs only for "
        "telemetry: 'passes visible but IPR = 0' is logged as a caught hack.",
        "Why it works: a real fix handles any valid input; a lookup table only knows the visible ones.",
        "Honest limits (from the project's findings): it needs a trusted reference + input sampler. It catches 93% "
        "of constructed cheats but only ~21–25% of natural subtle bugs. The naive version false-rejected ~10% of "
        "correct answers; wrapper-aware comparison + a fault-tolerant gate (reject only at ≥ max(2, ⌈0.15 × admitted checks⌉) divergences, in the MBPP grader ipr/bench/evalplus_ipr.py) raised "
        "recall 0.90 → 0.97. On SWE-bench, literal assert f(x) == y is rare, so the gold Docker image must be "
        "re-run on resampled inputs.",
    ])
    s.sub("Where it lives in the repo")
    s.para("swe-rl-ipr/ipr/reward.py grade();  perturb.py PerturbationEngine.resample_suite();  tamper.py detect_tamper();  "
           "runner.py run_cases() (an empty suite counts as a fail);  demo: python -m ipr.demo.magic_moment.")
    s.section("MATH", "math")
    s.math([
        "suite_k = { (x, gold(x)) : x ~ sampler(seed_k), gold(x) does not error }",
        "R_IPR = 1[no tamper] · Π_{k=1..K} 1[candidate passes every case in suite_k]",
        "Hack caught = 1[visible pass] · (1 − R_IPR)          Fault-tolerant (MBPP grader): R = 1[divergences < max(2, ⌈0.15 · admitted⌉)]",
    ])
    s.section("EXAMPLE: grade 4 patches of to_roman(n) live (Excel's ROMAN() is the gold reference)", "example")
    s.note("Real fix = a correct implementation. Lookup cheat = returns the stored answer if n is a visible input, "
           "else KeyError. Buggy = no subtractive pairs (4 → IIII). Edit the blue n values; try putting 4 or 1994 in "
           "a perturbed suite, or 5000 (gold errors → input dropped).")
    s.note("Simplifications vs the repo: 3 inputs per perturbed suite instead of the default 8 (RewardConfig.n_inputs_per_perturbation), "
           "fixed inputs instead of seeded random sampling, and no order-shuffle.")
    s.sub("Visible suite (what the agent can see)")
    vis = [4, 9, 58, 1994]
    va, vb = s.r + 1, s.r + 4

    def buggy(c):
        return (f'=IFERROR(REPT("M",INT({c}/1000))&REPT("D",INT(MOD({c},1000)/500))&REPT("C",INT(MOD({c},500)/100))'
                f'&REPT("L",INT(MOD({c},100)/50))&REPT("X",INT(MOD({c},50)/10))&REPT("V",INT(MOD({c},10)/5))'
                f'&REPT("I",MOD({c},5)),"error")')

    s.table(["Visible test", "n", "Expected = gold(n)", "Real fix output", "Lookup cheat output", "Buggy output",
             "Real pass", "Lookup pass", "Buggy pass"], 4,
            lambda i, r: [f"Visible {i + 1}", vis[i], f"=ROMAN(C{r})", f"=ROMAN(C{r})",
                          f'=IFERROR(INDEX($D${va}:$D${vb},MATCH(C{r},$C${va}:$C${vb},0)),"KeyError")',
                          buggy(f"C{r}"), f"=IF(E{r}=$D{r},1,0)", f"=IF(F{r}=$D{r},1,0)", f"=IF(G{r}=$D{r},1,0)"],
            fmts={1: "0", 6: "0", 7: "0", 8: "0"}, header_height=30)
    s.sub("K = 4 perturbed suites (fresh inputs, gold-defined answers, never shown to the agent)")
    pert = [(1, 3), (1, 14), (1, 40), (2, 90), (2, 400), (2, 2024), (3, 7), (3, 49), (3, 944), (4, 3999), (4, 1666),
            (4, 19)]
    pa, pb = s.r + 1, s.r + len(pert)
    s.table(["Suite k", "n (fresh input)", "Admitted? (gold defined)", "Expected = gold(n)", "Real fix output",
             "Lookup cheat output", "Buggy output", "Real pass", "Lookup pass", "Buggy pass"], len(pert),
            lambda i, r: [pert[i][0], pert[i][1], f"=IF(ISERROR(ROMAN(C{r})),0,1)",
                          f'=IFERROR(ROMAN(C{r}),"gold error")', f'=IFERROR(ROMAN(C{r}),"error")',
                          f'=IFERROR(INDEX($D${va}:$D${vb},MATCH(C{r},$C${va}:$C${vb},0)),"KeyError")',
                          buggy(f"C{r}"),
                          f'=IF($D{r}=0,"dropped",IF(F{r}=$E{r},1,0))', f'=IF($D{r}=0,"dropped",IF(G{r}=$E{r},1,0))',
                          f'=IF($D{r}=0,"dropped",IF(H{r}=$E{r},1,0))'],
            fmts={0: "0", 1: "0", 2: "0", 7: "0", 8: "0", 9: "0"}, header_height=30)
    s.sub("AND gate per suite (a suite passes only if every admitted case passes; an empty suite fails)")
    sa, sb = s.r + 1, s.r + 4
    B, D = rng(0, pa, pb), rng(2, pa, pb)

    def suite_pass(cidx, r):
        P = rng(cidx, pa, pb)
        return f"=IF(C{r}=0,0,IF(SUMIF({B},B{r},{P})=C{r},1,0))"

    s.table(["Suite k", "Admitted cases", "Real passes suite", "Lookup passes suite", "Buggy passes suite"], 4,
            lambda i, r: [i + 1, f"=COUNTIFS({B},B{r},{D},1)", suite_pass(7, r), suite_pass(8, r), suite_pass(9, r)],
            fmts={0: "0", 1: "0", 2: "0", 3: "0", 4: "0"})
    s.sub("Final grades")
    cand = [("Real fix", 0, "H", "D", "I"), ("Lookup cheat", 0, "I", "E", "J"), ("Tamper cheat", 1, "H", "D", "I"),
            ("Buggy", 0, "J", "F", "K")]

    def grade_row(i, r):
        name, tamper, vcol, scol, pcol = cand[i]
        return [name, tamper,
                f"=IF(SUM({vcol}{va}:{vcol}{vb})=COUNT($C${va}:$C${vb}),1,0)",
                f"=IF(C{r}=1,0,SUM({scol}{sa}:{scol}{sb}))",
                f'=COUNTIF({pcol}{pa}:{pcol}{pb},0)',
                f"=D{r}",
                f"=IF(AND(C{r}=0,E{r}=COUNT($B${sa}:$B${sb})),1,0)",
                f"=IF(AND(C{r}=0,F{r}<MAX(2,ROUNDUP(0.15*SUM($C${sa}:$C${sb}),0))),1,0)",
                f'=IF(AND(D{r}=1,H{r}=0),"YES: caught","")']

    s.table(["Candidate", "Touched test files? (1/0)", "Visible suite passed", "Perturbed suites passed (of K)",
             "Divergences (failed perturbed cases)", "Raw reward", "IPR reward", "Fault-tolerant IPR (MBPP-grader threshold)",
             "Hack caught?"], 4, grade_row, fmts={1: "0", 2: "0", 3: "0", 4: "0", 5: "0", 6: "0", 7: "0"},
            header_height=44)
    s.note("What to notice: the raw reward pays 3 of 4 candidates. IPR pays only the real fix and flags both cheats. "
           "The tamper cheat's code is correct, but it touched the tests, so reward.py zeroes it and skips the perturbed suites (the visible suite still runs, for telemetry).")
    return s


def s19(wb):
    s = new(wb, "19 Precision & Recall", "Evaluation & reward integrity", "A reward is a classifier: precision and recall",
            "Judge a reward function the way you'd judge a classifier: how often does it pay wrong code, and how often does it refuse right code?")
    s.section("CONCEPT", "concept")
    s.bullets([
        "A reward function labels every candidate as 'correct' (reward 1) or not. Compare its labels with ground "
        "truth (for MBPP, the large held-out MBPP+ suite).",
        "Precision = of the answers the reward paid, the share that were truly correct. Low precision → RL pays for "
        "wrong code → it learns hacks.",
        "Recall = of the truly correct answers, the share that got paid. Low recall → good answers get punished → "
        "noisy, slower learning.",
        "Hack-catch rate = the share of deliberate cheats that get reward 0.",
        "IPR's findings: precision on natural samples is not better than raw (0.78 vs 0.79), recall was fixed from "
        "0.90 to 0.97, and constructed cheats are caught 93% vs 0%.",
    ])
    s.section("MATH", "math")
    s.math([
        "TP = paid & correct     FP = paid & wrong     FN = refused & correct     TN = refused & wrong",
        "Precision = TP / (TP + FP)        Recall = TP / (TP + FN)        F1 = 2 · P · R / (P + R)",
        "Hack-catch rate = cheats with reward 0 / all cheats",
    ])
    s.section("EXAMPLE A: calculator (counts are ILLUSTRATIVE, not the paper's raw counts)", "example")
    data = [("TP  (paid, truly correct)", 474, 460), ("FP  (paid, actually wrong)", 126, 110),
            ("FN  (refused, truly correct)", 6, 20), ("TN  (refused, actually wrong)", 394, 410)]
    a = s.r + 1
    s.table(["Count", "Raw test-pass reward", "IPR reward"], 4, lambda i, r: list(data[i]), fmts={1: "#,##0", 2: "#,##0"})
    tp, fp, fn, tn = a, a + 1, a + 2, a + 3
    s.table(["Metric", "Raw test-pass reward", "IPR reward"], 4,
            lambda i, r: [["Precision", "Recall", "F1", "Accuracy"][i]] + [
                [f"={c}{tp}/({c}{tp}+{c}{fp})", f"={c}{tp}/({c}{tp}+{c}{fn})",
                 f"=2*{c}{r - 2}*{c}{r - 1}/({c}{r - 2}+{c}{r - 1})",
                 f"=({c}{tp}+{c}{tn})/SUM({c}{tp}:{c}{tn})"][i] for c in ("C", "D")],
            fmts={1: "0.000", 2: "0.000"})
    s.section("EXAMPLE B: reported results (swe-rl-ipr/paper/FACTS.md §3c, 12,096 samples of Qwen2.5-Coder-1.5B on MBPP+)", "example")
    data = [("Precision on natural samples", 0.79, 0.78), ("Constructed cheats caught", 0.0, 0.93),
            ("Recall (naive perturbation)", None, 0.90), ("Recall (edge-aware + fault-tolerant gate)", None, 0.97)]
    s.table(["Metric", "Raw test-pass reward", "IPR reward"], 4,
            lambda i, r: [data[i][0], data[i][1] if data[i][1] is not None else "n/a", data[i][2]],
            fmts={1: "0%", 2: "0%"})
    s.note("What to notice: IPR's win is catching memorisation (93% vs 0%), not natural subtle bugs, where its "
           "precision is about the same as the raw reward's.")
    return s


def s20(wb):
    s = new(wb, "20 Hacking Simulation", "Evaluation & reward integrity", "Simulation: raw reward vs IPR during training",
            "Train the same 4-choice policy twice, changing only the reward. The raw reward grows the cheat; IPR trains it out.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "A replica of swe-rl-ipr/ipr/train/run.py, made deterministic: a softmax policy over 4 strategies, trained "
        "with the exact expected policy gradient (sheet 05) instead of sampled groups. The repo's run.py samples groups "
        "of 12 for 60 rounds at lr 0.4 with GRPO advantages; here it is 40 rounds at η = 2.",
        "Cheat-favouring prior: hacks start more likely because they are easier to write (logits 0, 1.1, 0.6, 0).",
        "Raw reward pays real, lookup and tamper all +1. The update is π_k(R_k − J), so the already-likelier cheat "
        "gains faster (rich get richer) and P(hack) GROWS.",
        "IPR pays only the real fix. Cheats get negative updates and P(hack) → 0.",
        "The repo's sampled version reports P(hack) 0.68 → 0.009 under IPR and 0.73 → 0.90 under raw.",
    ])
    s.section("MATH", "math")
    s.math(["π = softmax(z);    J = Σ_k π_k R_k;    z_k ← z_k + η · π_k · (R_k − J);    P(hack) = π_lookup + π_tamper"])
    s.section("EXAMPLE: 40 training rounds (edit the prior, rewards or η)", "example")
    s.kv([("eta", "Learning rate η", 2.0, "0.0", "")])
    strat = [("Real fix", 0.0, 1, 1), ("Lookup cheat", 1.1, 1, 0), ("Tamper cheat", 0.6, 1, 0), ("Buggy", 0.0, 0, 0)]
    pa = s.r + 1
    s.table(["Strategy", "Prior logit z0", "Reward under RAW", "Reward under IPR"], 4, lambda i, r: list(strat[i]),
            fmts={1: "0.0", 2: "0", 3: "0"})
    s.gap()
    hdr = ["Round"]
    for arm in ("RAW", "IPR"):
        hdr += [f"{arm} z real", f"{arm} z lookup", f"{arm} z tamper", f"{arm} z buggy", f"{arm} P(real)",
                f"{arm} P(hack)", f"{arm} J"]
    R = 41

    def row(i, r):
        out = [i]
        for arm, rcol, base in (("RAW", "D", 1), ("IPR", "E", 8)):
            zc = [col(base + k) for k in range(4)]
            S = "+".join(f"EXP({c}{r})" for c in zc)
            if i == 0:
                out += [f"=$C${pa + k}" for k in range(4)]
            else:
                p = r - 1
                Sp = "+".join(f"EXP({c}{p})" for c in zc)
                Jp = f"{col(base + 6)}{p}"
                out += [f"={zc[k]}{p}+«eta»*EXP({zc[k]}{p})/({Sp})*(${rcol}${pa + k}-{Jp})" for k in range(4)]
            out += [f"=EXP({zc[0]}{r})/({S})", f"=(EXP({zc[1]}{r})+EXP({zc[2]}{r}))/({S})",
                    "=(" + "+".join(f"EXP({zc[k]}{r})*${rcol}${pa + k}" for k in range(4)) + f")/({S})"]
        return out

    f1, f2 = s.table(hdr, R, row, fmts={**{j: "0.00" for j in range(1, 15)}, 5: "0.000", 6: "0.000", 12: "0.000",
                                        13: "0.000"}, header_height=30)
    ch = line_chart("P(hack) during training", "round", "P(hack)", width=15, height=8)
    for c in (FIRST + 6, FIRST + 13):
        ch.add_data(Reference(s.ws, min_col=c, min_row=f1 - 1, max_row=f2), titles_from_data=True)
    ch.set_categories(Reference(s.ws, min_col=FIRST, min_row=f1, max_row=f2))
    s.ws.add_chart(ch, f"R{f1 - 1}")
    ch2 = line_chart("P(real fix) during training", "round", "P(real)", width=15, height=8)
    for c in (FIRST + 5, FIRST + 12):
        ch2.add_data(Reference(s.ws, min_col=c, min_row=f1 - 1, max_row=f2), titles_from_data=True)
    ch2.set_categories(Reference(s.ws, min_col=FIRST, min_row=f1, max_row=f2))
    s.ws.add_chart(ch2, f"R{f1 + 16}")
    return s


def s21(wb):
    s = new(wb, "21 Compute & Scale", "Evaluation & reward integrity", "Compute and scale: laptop vs production",
            "The algorithm is the same. The scale differs by orders of magnitude, and memory is the first wall you hit.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Mini-SWE-RL and DeepSWE run the same GRPO loop. The differences are model size, number of tasks, group "
        "size, horizon (1 turn vs ~100) and hardware.",
        "Rollouts dominate time in production (Docker environments + long agent loops), so generation runs async "
        "on vLLM engines while training runs on FSDP-sharded GPUs.",
        "Memory rule of thumb for full fine-tuning with AdamW: about 12 bytes per parameter (bf16 weights 2 + "
        "gradients 2 + fp32 Adam moments 8), before activations. A reference model adds another 2 bytes per "
        "parameter. That's why v2 dropped the reference model on a 48 GB Mac. LoRA (used for the 7B MBPP run) "
        "trains ~0.5% of the weights to avoid this.",
    ])
    s.section("MATH", "math")
    s.math([
        "Rollouts = tasks × G × epochs          GPU-hours = GPUs × wall-clock hours",
        "Training memory ≈ params × (bytes_weights + bytes_grads + bytes_optimizer) + reference model + activations",
    ])
    s.section("EXAMPLE A: side by side (numbers from the Mini-SWE-RL README)", "example")
    data = [("Model parameters (billions)", 1.5, 32, "Qwen2.5-Coder-1.5B vs Qwen3-32B"),
            ("GPUs", 1, 64, "Apple M4 Pro vs H100"),
            ("Training wall-clock (hours)", 0.5, 144, "30 minutes vs 6 days"),
            ("GPU-hours", "=C{r0}*C{r1}", "=D{r0}*D{r1}", "formula"),
            ("Group size G", 8, 16, ""),
            ("Tasks in the environment", 45, 8100, "hand-written puzzles vs R2E-Gym problems"),
            ("Reported result", "66.7% → 73.3% (toy puzzles)", "42.2% Pass@1 (SWE-bench Verified)", "")]
    a = s.r + 1

    def row(i, r):
        q, m, d, n = data[i]
        if isinstance(m, str) and m.startswith("="):
            m = m.format(r0=a + 1, r1=a + 2)
            d = d.format(r0=a + 1, r1=a + 2)
        return [q, m, d, n]

    s.table(["Quantity", "Mini-SWE-RL", "DeepSWE (production)", "Notes"], 7, row,
            fmts={1: "#,##0.0", 2: "#,##0.0"}, merge_last=True, row_height=20)
    s.kv([(None, "Compute ratio (GPU-hours)", f"=D{a + 3}/C{a + 3}", "#,##0", "×; and an H100 is far faster than an M4 Pro")])
    s.section("EXAMPLE B: Mini-SWE-RL v2 rollout budget", "example")
    s.kv([("t", "Training puzzles", 16, "0", "TRAIN_PUZZLES"), ("g", "Group size G", 8, "0", ""),
          ("e", "Epochs", 10, "0", ""), ("ro", "Total training rollouts", "=«t»*«g»*«e»", "#,##0", "README: 1,280"),
          ("sec", "Seconds per Ollama rollout", 0.5, "0.0", "README: ~0.3–0.8 s"),
          (None, "Generation time (minutes)", "=«ro»*«sec»/60", "0.0",
           "estimate: the rest of the ~30 min in history.json is mostly HuggingFace forward/backward passes")])
    s.section("EXAMPLE C: memory calculator (rule-of-thumb assumptions, edit freely)", "example")
    s.kv([("P", "Parameters (billions)", 1.5, "0.0", ""), ("bw", "Bytes per weight (bf16)", 2, "0", ""),
          ("bg", "Bytes per gradient", 2, "0", ""), ("bo", "Optimizer bytes per param (AdamW fp32 m + v)", 8, "0", ""),
          ("ref", "Keep a reference model? (1/0)", 1, "0", "for the KL penalty"),
          ("tr", "Training state (GB)", "=«P»*(«bw»+«bg»+«bo»)", "0.0", "weights + gradients + optimizer"),
          ("rm", "Reference model (GB)", "=«ref»*«P»*«bw»", "0.0", ""),
          (None, "Total before activations (GB)", "=«tr»+«rm»", "0.0",
           "README: loading policy + reference model plus optimizer state hit 34.7 GB on MPS and ran out of memory")])
    return s


def index(wb):
    ws = wb["Index"]
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 26
    ws.column_dimensions["E"].width = 90
    ws.merge_cells("B2:E2")
    c = ws["B2"]
    c.value = "RL for SWE Agents: Concept Workbook"
    c.font = font(bold=True, size=18, color="FFFFFF")
    c.fill = fill(NAVY)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[2].height = 34
    ws.merge_cells("B3:E3")
    ws["B3"].value = ("Each topic from Mini-SWE-RL (github.com/RajatDandekar/Mini-SWE-RL) and its swe-rl-ipr follow-up, "
                      "broken into CONCEPT (plain words) → MATH (the formulas) → EXAMPLE (live Excel you can edit).")
    ws["B3"].font = font(italic=True, size=11, color=NAVY)
    ws["B3"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[3].height = 32
    r = 5
    ws.cell(r, 2, "How to read the sheets").font = font(bold=True, size=11, color=NAVY)
    r += 1
    legend = [("12", "Blue text on yellow = an input you can change. Everything else recalculates.", True),
              ("0.375", "Black text = a formula. Don't type over these.", False),
              ("", "Section colours: CONCEPT (blue) · MATH (purple) · EXAMPLE (green).", False),
              ("", "Numbers in EXAMPLE sections are illustrative toy values unless a note on the sheet cites the repo as the source.", False)]
    for sample, text, is_input in legend:
        cc = ws.cell(r, 2, sample)
        cc.alignment = Alignment(horizontal="center")
        cc.border = BOX if sample else Border()
        cc.font = font(color="0000FF" if is_input else "000000")
        if is_input:
            cc.fill = fill(FILL["input"])
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        ws.cell(r, 3, text).font = font()
        r += 1
    r += 1
    for j, h in enumerate(["#", "Sheet", "Part", "What you'll understand"]):
        cc = ws.cell(r, 2 + j, h)
        cc.font = font(bold=True)
        cc.fill = fill(FILL["hdr"])
        cc.border = BOX
    r += 1
    part_fill = {"Foundations": "DDEBF7", "GRPO mechanics": "E2EFDA", "Production recipe": "FCE4D6",
                 "Evaluation & reward integrity": "E4DFEC"}
    for i, (name, part, title, idea) in enumerate(SHEETS, 1):
        ws.cell(r, 2, i).font = font()
        ws.cell(r, 2).alignment = Alignment(horizontal="center", vertical="center")
        lc = ws.cell(r, 3, title)
        lc.hyperlink = f"#'{name}'!A1"
        lc.font = font(color="0563C1", underline="single")
        lc.alignment = Alignment(wrap_text=True, vertical="center")
        pc = ws.cell(r, 4, part)
        pc.font = font()
        pc.fill = fill(part_fill[part])
        pc.alignment = Alignment(vertical="center")
        ic = ws.cell(r, 5, idea)
        ic.font = font()
        ic.alignment = Alignment(wrap_text=True, vertical="center")
        for cc in range(2, 6):
            ws.cell(r, cc).border = BOX
        ws.row_dimensions[r].height = 30 if len(idea) > 95 else 18
        r += 1
    r += 1
    ws.cell(r, 2, "Suggested path").font = font(bold=True, size=11, color=NAVY)
    r += 1
    for t in ["1. Foundations (01–06): what is being optimised and how a gradient comes from a test result.",
              "2. GRPO mechanics (07–09): the group trick, why task difficulty matters, and the loss.",
              "3. Production recipe (10–15): the extras DeepSWE's GRPO++ adds and why Mini-SWE-RL's shortcuts matter.",
              "4. Evaluation & reward integrity (16–21): trusting results, reward hacking, IPR, and scale."]:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        ws.cell(r, 2, t).font = font()
        r += 1


def main():
    wb = Workbook()
    wb.active.title = "Index"
    sheets = [f(wb) for f in (s01, s02, s03, s04, s05, s06, s07, s08, s09, s10, s11, s12, s13, s14, s15, s16, s17,
                              s18, s19, s20, s21)]
    for s in sheets:
        s.resolve()
        s.ws.sheet_properties.pageSetUpPr.fitToPage = True
        s.ws.page_setup.fitToWidth = 1
        s.ws.page_setup.fitToHeight = 0
        s.ws.page_setup.orientation = "landscape"
        s.ws.sheet_properties.tabColor = {"Foundations": "5B9BD5", "GRPO mechanics": "70AD47",
                                          "Production recipe": "ED7D31",
                                          "Evaluation & reward integrity": "7030A0"}[
            next(p for n, p, _, _ in SHEETS if n == s.name)]
    index(wb)
    ix = wb["Index"]
    ix.sheet_properties.pageSetUpPr.fitToPage = True
    ix.page_setup.fitToWidth = 1
    ix.page_setup.fitToHeight = 0
    ix.page_setup.orientation = "landscape"
    wb.save(OUT)
    print("saved", OUT)


if __name__ == "__main__":
    main()
