"""Build RL_Fundamentals.xlsx — reinforcement learning from scratch, one concept per sheet.

Every sheet follows Concept → Math → Example, and every example is live Excel formulas built on two
running toy problems: a 7-cell corridor world and a 3-armed bandit. Change a blue cell on a yellow
background and the sheet recomputes.

Reuses the layout engine from ../rl-swe-agents/build_workbook.py. Run:  python build_rl_fundamentals.py
"""

import random
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "rl-swe-agents"))

from build_workbook import (  # noqa: E402  (shared layout engine)
    BOX, FILL, FIRST, NAVY, Sheet, col, fill, font, line_chart,
)
from openpyxl import Workbook  # noqa: E402
from openpyxl.chart import Reference  # noqa: E402
from openpyxl.styles import Alignment, Border  # noqa: E402

OUT = "RL_Fundamentals.xlsx"
PARTS = {
    "Foundations": ("DDEBF7", "5B9BD5"),
    "Bandits & exploration": ("E2EFDA", "70AD47"),
    "Planning with a model": ("FCE4D6", "ED7D31"),
    "Learning from experience": ("E4DFEC", "7030A0"),
    "Policies & modern RL": ("DEEAF1", "1F7A8C"),
}
SHEETS = []  # (name, part, title, idea)
SHEET_OBJS = []

# Running example 1: the corridor. Cells 1..5 are normal, 0 and 6 are exits (episode ends).
# Entering cell 0 pays +1 (small, close), entering cell 6 pays +10 (big). Start in cell 3.
R_LEFT, R_RIGHT, START = 1, 10, 3
# Running example 2: a 3-armed bandit with hidden win rates.
MU = [0.2, 0.5, 0.8]


def new(wb, name, part, title, idea, widths=None):
    SHEETS.append((name, part, title, idea))
    sheet = Sheet(wb, name, title, idea, widths)
    SHEET_OBJS.append(sheet)
    return sheet


def corridor_strip(s, cells, label="Corridor", fmt=None, height=30):
    """Draw the 7-cell corridor as coloured boxes in columns C..I. cells: 7 values or formulas."""
    kinds = ["exitS", "cell", "cell", "start", "cell", "cell", "exitB"]
    colours = {"exitS": "FCE4D6", "exitB": "C6EFCE", "start": "DDEBF7", "cell": "FFFFFF"}
    r = s.r
    s.ws.cell(r, FIRST, label).font = font(bold=True)
    s.ws.cell(r, FIRST).alignment = Alignment(vertical="center")
    for j, v in enumerate(cells):
        c = s.ws.cell(r, FIRST + 1 + j, v)
        c.border = BOX
        c.fill = fill(colours[kinds[j]])
        c.font = font(bold=kinds[j] != "cell", size=11)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if fmt:
            c.number_format = fmt
    s.ws.row_dimensions[r].height = height
    s.r += 1


def corridor_kv(s, gamma=0.9):
    s.kv([("rL", "Reward for the small exit (cell 0)", R_LEFT, "0", "close to the start"),
          ("rR", "Reward for the big exit (cell 6)", R_RIGHT, "0", "further away"),
          ("g", "Discount factor γ", gamma, "0.00", "sheet 02")])


def random_walk(seed, n_episodes, p_right=0.5):
    """Episodes of a random walk in the corridor from START: lists of (s, a, r, s')."""
    rng = random.Random(seed)
    episodes = []
    for _ in range(n_episodes):
        s, tr = START, []
        while s not in (0, 6):
            a = 1 if rng.random() < p_right else -1
            s2 = s + a
            tr.append((s, a, R_RIGHT if s2 == 6 else R_LEFT if s2 == 0 else 0, s2))
            s = s2
        episodes.append(tr)
    return episodes


def reward_into(cell):
    return f"IF({cell}=6,«rR»,IF({cell}=0,«rL»,0))"


# =============================================================================
def s01(wb):
    s = new(wb, "01 The RL Loop", "Foundations", "The reinforcement learning loop",
            "An agent acts, the environment answers with a new situation and a reward, and the agent learns to collect as much reward as possible.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Reinforcement learning (RL) is learning by trial and error. Nobody shows the agent the right move; it only "
        "gets a score (reward) after acting, and must work out which behaviour leads to high scores.",
        "Agent = the learner and decision maker. Environment = everything it interacts with.",
        "State s = the situation the agent is in. Action a = what it chooses to do. Reward r = a number the "
        "environment returns after each action (good = high).",
        "Episode = one run from a start to an end (for example, one game). The goal is the largest TOTAL reward over "
        "the episode, not the largest single reward.",
        "How it differs from supervised learning: there are no labelled right answers, feedback can be delayed "
        "(the reward may come many steps after the action that earned it), and the agent's own choices decide "
        "what data it sees.",
        "Everyday examples: game-playing agents, robot control, recommendation systems, and training language "
        "models to give better answers (sheet 18).",
    ])
    s.sub("Running example used throughout this workbook: the corridor")
    s.para("Seven cells in a row. The agent starts in cell 3 and can move left (−1) or right (+1). Cell 0 is a small "
           "exit worth +1 (close); cell 6 is a big exit worth +10 (further). Reaching either exit ends the episode. "
           "Every other move pays 0.")
    corridor_strip(s, ["EXIT\n+1", "1", "2", "3\nSTART", "4", "5", "EXIT\n+10"], height=34)
    s.section("MATH", "math")
    s.math([
        "At each step t:  observe state s_t  →  choose action a_t  →  receive reward r_t+1 and next state s_t+1",
        "Trajectory (one episode):  s_0, a_0, r_1, s_1, a_1, r_2, s_2, …, s_T   (T = the step the episode ends)",
        "Total reward of the episode = r_1 + r_2 + … + r_T     (sheet 02 adds discounting)",
    ])
    s.legend([("s_t", "state at step t"), ("a_t", "action taken at step t"),
              ("r_t+1", "reward received after that action"), ("T", "final step of the episode")])
    s.section("EXAMPLE: one episode in the corridor (edit the moves)", "example")
    s.kv([("start", "Start cell", START, "0", "")])
    s.kv([("rL", "Reward, small exit", R_LEFT, "0", ""), ("rR", "Reward, big exit", R_RIGHT, "0", "")])
    moves = [1, -1, 1, 1, 1, 1, -1]
    a, b = s.r + 1, s.r + len(moves)
    s.table(["Step t", "State s_t", "Move a_t (+1 right, −1 left)", "Next state s_t+1", "Reward r_t+1",
             "Episode over?"], len(moves),
            lambda i, r: [i, "=«start»" if i == 0 else f"=E{r - 1}", moves[i],
                          f"=IF(OR(C{r}=0,C{r}=6),C{r},C{r}+D{r})",
                          f"=IF(OR(C{r}=0,C{r}=6),0,{reward_into(f'E{r}')})", f"=OR(E{r}=0,E{r}=6)"],
            fmts={0: "0", 1: "0", 2: "0", 3: "0", 4: "0"}, header_height=30)
    s.gap()
    s.kv([(None, "Total reward", f"=SUM(F{a}:F{b})", "0", ""),
          (None, "Steps until an exit", f'=IFERROR(MATCH(TRUE,G{a}:G{b},0),"not finished")', "0",
           "moves after the exit are ignored")])
    s.note("What to notice: the agent wandered (right, back left, then right three times) and still earned +10. "
           "Change the first three moves to −1 and it takes the +1 exit after 3 steps instead. RL is about learning "
           "which sequence of moves is best.")


def s02(wb):
    s = new(wb, "02 Return & Discount", "Foundations", "Return and the discount factor γ",
            "The return is the total future reward, with later rewards counted for less. γ decides how far ahead the agent cares.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Return G = the sum of all rewards from now until the episode ends. The agent tries to maximise the "
        "return, not the next reward.",
        "Discounting: a reward k steps in the future is multiplied by γ^k, with γ between 0 and 1. This expresses "
        "'sooner is better', models uncertainty about the future, and keeps sums finite in endless tasks.",
        "γ = 0: only the next reward matters (short-sighted). γ close to 1: far-future rewards count almost fully "
        "(far-sighted). A rough planning horizon is 1 / (1 − γ) steps.",
        "The return can be computed backwards in one pass: G_t = r_t+1 + γ · G_t+1. This recursive form is the "
        "seed of every value-learning method later in the workbook.",
    ])
    s.section("MATH", "math")
    s.math([
        "G_t = r_t+1 + γ·r_t+2 + γ²·r_t+3 + …  =  Σ_k γ^k · r_t+k+1",
        "Recursive form:  G_t = r_t+1 + γ · G_t+1          (G_T = 0 at the end)",
        "Effective horizon ≈ 1 / (1 − γ)",
    ])
    s.section("EXAMPLE A: return of a reward sequence", "example")
    s.kv([("g", "Discount factor γ", 0.9, "0.00", "try 0.5 and 0.99")])
    rewards = [0, 0, 0, 0, 0, 10]
    a, b = s.r + 1, s.r + len(rewards)
    s.table(["Step t", "Reward r_t+1", "Discount γ^t", "Discounted reward", "Return G_t (computed backwards)"],
            len(rewards),
            lambda i, r: [i, rewards[i], f"=«g»^B{r}", f"=C{r}*D{r}",
                          f"=C{r}+«g»*F{r + 1}" if i < len(rewards) - 1 else f"=C{r}"],
            fmts={0: "0", 1: "0", 2: "0.000", 3: "0.000", 4: "0.000"}, header_height=30)
    s.gap()
    s.kv([(None, "G_0 = sum of discounted rewards", f"=SUM(E{a}:E{b})", "0.000", "matches the top of column F"),
          (None, "Effective horizon (steps)", '=IF(«g»=1,"unlimited",1/(1-«g»))', "0.0", "")])
    s.section("EXAMPLE B: +1 now, or +10 after 5 more steps?", "example")
    s.kv([("small", "Small reward, received immediately", 1, "0", ""),
          ("big", "Big reward", 10, "0", ""), ("delay", "Delay before the big reward (steps)", 5, "0", ""),
          (None, "Break-even γ", "=(«small»/«big»)^(1/«delay»)", "0.000",
           "above this γ, waiting is worth more")])
    gammas = [0, 0.3, 0.5, 0.6, 0.63, 0.7, 0.8, 0.9, 0.99, 1]
    f1, f2 = s.table(["γ", "Return if you take +1 now", "Return if you wait for +10", "Better choice"], len(gammas),
                     lambda i, r: [gammas[i], "=«small»", f"=«big»*B{r}^«delay»",
                                   f'=IF(D{r}>C{r},"wait for the big reward","take the small reward now")'],
                     fmts={0: "0.00", 1: "0.000", 2: "0.000"})
    ch = line_chart("Return of each choice vs γ", "γ", "return", width=14, height=8)
    ch.add_data(Reference(s.ws, min_col=FIRST + 1, max_col=FIRST + 2, min_row=f1 - 1, max_row=f2),
                titles_from_data=True)
    ch.set_categories(Reference(s.ws, min_col=FIRST, min_row=f1, max_row=f2))
    s.ws.add_chart(ch, f"G{f1 - 1}")
    s.note("What to notice: the same rewards lead to opposite decisions depending on γ. γ is part of the problem "
           "definition, so choose it on purpose.")


def s03(wb):
    s = new(wb, "03 MDP", "Foundations", "Markov decision process (MDP)",
            "An MDP is the formal recipe for an RL problem: states, actions, what happens after each action (possibly at random), and the rewards.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Markov property: the next state and reward depend only on the CURRENT state and action, not on the "
        "history of how you got there. The state must contain everything that matters.",
        "An MDP has five parts: states S, actions A, transition probabilities P(s' | s, a), rewards R, and the "
        "discount γ.",
        "Transitions can be random: the same action from the same state can lead to different outcomes "
        "(a slippery floor, an opponent, luck). For each (state, action) pair the outcome probabilities must sum to 1.",
        "If you KNOW P and R you can plan with them (sheets 08–10). If you don't, you must learn from experience "
        "(sheets 11–14). That split, model-based vs model-free, organises most of RL.",
    ])
    s.section("MATH", "math")
    s.math([
        "MDP = (S, A, P, R, γ)",
        "P(s' | s, a) = probability of landing in s' after taking a in s;   Σ_s' P(s' | s, a) = 1",
        "Expected immediate reward:  r(s, a) = Σ_s' P(s' | s, a) · R(s, a, s')",
    ])
    s.section("EXAMPLE: a recycling robot", "example")
    s.note("Adapted from Sutton & Barto, Reinforcement Learning: An Introduction (2nd ed.), Example 3.3. The numbers "
           "here are illustrative. The robot's battery is High or Low. Searching finds cans (reward) but drains the "
           "battery; if a Low battery runs out the robot must be rescued (penalty) and is recharged to High.")
    s.kv([("alpha", "P(battery stays High | search)", 0.7, "0.00", ""),
          ("beta", "P(battery stays Low | search)", 0.6, "0.00", "otherwise it runs flat"),
          ("rs", "Reward for searching", 2, "0", "cans found"), ("rw", "Reward for waiting", 1, "0", "fewer cans"),
          ("rescue", "Reward when rescued", -3, "0", "battery ran flat")])
    rows = [("High", "search", "High", "=«alpha»", "=«rs»"), ("High", "search", "Low", "=1-«alpha»", "=«rs»"),
            ("High", "wait", "High", 1, "=«rw»"), ("Low", "search", "Low", "=«beta»", "=«rs»"),
            ("Low", "search", "High (rescued)", "=1-«beta»", "=«rescue»"), ("Low", "wait", "Low", 1, "=«rw»"),
            ("Low", "recharge", "High", 1, 0)]
    a, b = s.r + 1, s.r + len(rows)
    s.table(["State s", "Action a", "Next state s'", "P(s' | s, a)", "Reward R", "P × R"], len(rows),
            lambda i, r: list(rows[i]) + [f"=E{r}*F{r}"], fmts={3: "0.00", 4: "0", 5: "0.00"},
            label_cols=(0, 1, 2))
    S, A, P, PR = f"$B${a}:$B${b}", f"$C${a}:$C${b}", f"$E${a}:$E${b}", f"$G${a}:$G${b}"
    pairs = [("High", "search"), ("High", "wait"), ("Low", "search"), ("Low", "wait"), ("Low", "recharge")]
    s.sub("Summary per (state, action)")
    s.table(["State s", "Action a", "Σ P (must be 1)", "Expected reward r(s, a)"], len(pairs),
            lambda i, r: [pairs[i][0], pairs[i][1], f'=SUMIFS({P},{S},B{r},{A},C{r})',
                          f'=SUMIFS({PR},{S},B{r},{A},C{r})'],
            fmts={2: "0.00", 3: "0.00"}, label_cols=(0, 1))
    s.note("What to notice: with a Low battery, searching is worth 0 on average (cans vs rescue risk) and waiting is "
           "worth 1, so 'search' is only good when the battery is High. The best action depends on the state.")


def s04(wb):
    s = new(wb, "04 Policy", "Foundations", "Policy: the agent's decision rule",
            "A policy says what the agent does in each state, either always the same action or with probabilities. RL searches for the best policy.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Policy π = a rule mapping states to actions. It's the thing RL ultimately produces.",
        "Deterministic policy: one fixed action per state ('if Low, recharge'). Stochastic policy: a probability "
        "for each action ('if Low, search 30% of the time').",
        "Stochastic policies are useful for exploring (sheet 06) and are what policy-gradient methods learn "
        "(sheet 15). A language model is a stochastic policy over the next token (sheet 18).",
        "Each policy produces its own expected reward. Comparing policies by immediate reward alone is short-"
        "sighted; sheets 08–10 compare them by long-run value.",
    ])
    s.section("MATH", "math")
    s.math([
        "π(a | s) = probability of choosing action a in state s;   Σ_a π(a | s) = 1 for every state",
        "Expected immediate reward under π:   r_π(s) = Σ_a π(a | s) · r(s, a)",
    ])
    s.section("EXAMPLE: three policies for the recycling robot (r(s, a) from sheet 03)", "example")
    rows = [("High", "search", 2, 1, 1, 0.5), ("High", "wait", 1, 0, 0, 0.5), ("Low", "search", 0, 1, 0, 0.33),
            ("Low", "wait", 1, 0, 0, 0.33), ("Low", "recharge", 0, 0, 1, 0.34)]
    a, b = s.r + 1, s.r + len(rows)
    s.table(["State", "Action", "r(s, a)", "π: always search", "π: cautious", "π: random"], len(rows),
            lambda i, r: list(rows[i]), fmts={2: "0.00", 3: "0.00", 4: "0.00", 5: "0.00"}, label_cols=(0, 1))
    S, R = f"$B${a}:$B${b}", f"$D${a}:$D${b}"

    def summary(i, r):
        st = ["High", "Low"][i]
        checks = [f'=SUMIF({S},"{st}",{c}${a}:{c}${b})' for c in ("E", "F", "G")]
        values = [f'=SUMPRODUCT(({S}="{st}")*{R}*{c}${a}:{c}${b})' for c in ("E", "F", "G")]
        return [st] + checks + values

    s.sub("Per-state results")
    s.table(["State", "Σπ always search", "Σπ cautious", "Σπ random", "r_π always search", "r_π cautious",
             "r_π random"], 2, summary, fmts={j: "0.00" for j in range(1, 7)}, header_height=30)
    s.note("What to notice: on immediate reward alone, 'always search' and 'cautious' look identical (2 in High, 0 in "
           "Low), and 'random' even looks better in Low. But searching on a Low battery risks the −3 rescue, while "
           "'cautious' recharges to High where searching pays 2. Seeing that difference needs values over time, "
           "which is what sheets 08–10 compute.")


def s05(wb):
    s = new(wb, "05 Action Values", "Bandits & exploration", "Bandits and estimating action values",
            "With one state and several choices (a bandit), the value of an action is its average reward. Estimate it by averaging, updated one sample at a time.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Multi-armed bandit: one situation, k choices ('arms', like slot machines), each paying out with its own "
        "unknown average. The simplest RL problem: no states to move between, just choices.",
        "Action value q(a) = the true average reward of arm a. The agent doesn't know it, so it keeps an "
        "estimate Q(a) = the average of the rewards seen so far from that arm.",
        "Averages can be updated without storing every reward: NewEstimate = OldEstimate + StepSize × (Target − "
        "OldEstimate). This update pattern appears in almost every RL method (TD, Q-learning, actor-critic).",
        "With a constant step size α instead of 1/n, recent rewards count more. That's better when the arms' "
        "payouts change over time.",
    ])
    s.sub("Running example 2: a 3-armed bandit")
    s.para("Three slot machines with hidden win rates 0.2, 0.5 and 0.8 (sheets 05–07 and 15). Each pull pays 1 for a "
           "win and 0 otherwise.")
    s.section("MATH", "math")
    s.math([
        "Sample average:  Q_n = (R_1 + R_2 + … + R_n) / n",
        "Incremental form:  Q_n = Q_n−1 + (1/n) · (R_n − Q_n−1)",
        "Constant step size:  Q ← Q + α · (R − Q)          (recent rewards weigh more)",
    ])
    s.section("EXAMPLE: 12 pulls of one arm", "example")
    s.kv([("alpha", "Constant step size α", 0.2, "0.00", ""), ("q0", "Starting estimate Q0", 0, "0.00", "")])
    rewards = [1, 0, 0, 1, 1, 0, 1, 1, 0, 1, 1, 1]
    f1, f2 = s.table(["Pull n", "Reward R_n", "Sample average", "Incremental Q (1/n)", "Difference (check = 0)",
                      "Constant-α Q"], len(rewards),
                     lambda i, r: [i + 1, rewards[i], f"=AVERAGE($C${r - i}:C{r})",
                                   f"=«q0»+(C{r}-«q0»)/B{r}" if i == 0 else f"=E{r - 1}+(C{r}-E{r - 1})/B{r}",
                                   f"=E{r}-D{r}",
                                   f"=«q0»+«alpha»*(C{r}-«q0»)" if i == 0 else f"=G{r - 1}+«alpha»*(C{r}-G{r - 1})"],
                     fmts={0: "0", 1: "0", 2: "0.000", 3: "0.000", 4: "0.000", 5: "0.000"}, header_height=30)
    ch = line_chart("Estimate of the arm's value", "pull", "estimate", width=14, height=8)
    for c in (FIRST + 2, FIRST + 5):
        ch.add_data(Reference(s.ws, min_col=c, min_row=f1 - 1, max_row=f2), titles_from_data=True)
    ch.set_categories(Reference(s.ws, min_col=FIRST, min_row=f1, max_row=f2))
    s.ws.add_chart(ch, f"B{s.r + 1}")
    s.r += 17
    s.note("What to notice: the incremental update reproduces the average exactly without storing past rewards. "
           "The constant-α estimate reacts faster to the recent run of wins.")


def s06(wb):
    s = new(wb, "06 Epsilon-Greedy", "Bandits & exploration", "Exploration vs exploitation: ε-greedy",
            "Always picking the best-looking arm can lock you onto a bad one. Exploring at random a small fraction ε of the time fixes that.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Exploit = pick the arm with the highest estimate (greedy). Explore = try something else to learn more. "
        "You can't do both on the same pull: that's the exploration–exploitation dilemma.",
        "A purely greedy agent can get stuck: if its first choice pays a little, it never tries the others and "
        "never discovers they're better.",
        "ε-greedy: with probability ε pick a random arm, otherwise pick the greedy one. Simple and effective.",
        "Optimistic initial values: start every estimate high (e.g. Q0 = 1). Each disappointing arm then drops "
        "below the untried ones, so even a greedy agent tries everything early.",
    ])
    s.section("MATH", "math")
    s.math([
        "With probability ε:  a = random arm;   otherwise:  a = argmax_a Q(a)",
        "P(choose the greedy arm) = 1 − ε + ε/k;   P(choose each other arm) = ε/k",
        "After each pull:  N(a) ← N(a) + 1;   Q(a) ← Q(a) + (R − Q(a)) / N(a)        (sheet 05)",
    ])
    s.section("EXAMPLE: 100 pulls by two agents on the same random draws", "example")
    s.note("The u columns are pre-drawn random numbers (fixed so the sheet is reproducible). Both agents see the same "
           "draws; only ε differs. Try ε_A = 0 with Q0 = 1 to see optimistic starts fix the greedy agent.")
    s.kv([("epsA", "ε for agent A", 0.0, "0.00", "purely greedy"), ("epsB", "ε for agent B", 0.1, "0.00", ""),
          ("q0", "Starting estimate Q0 (both agents)", 0.0, "0.00", "try 1.0")])
    a, b = s.r + 1, s.r + 3
    s.table(["Arm", "True win rate μ (hidden)"], 3, lambda i, r: [f"Arm {i + 1}", MU[i]], fmts={1: "0.00"})
    s.names["mu"] = f"$C${a}:$C${b}"
    s.sub("Results after 100 pulls")
    s.kv([(None, "Agent A: average reward", "=«avgA»", "0.00", ""),
          (None, "Agent B: average reward", "=«avgB»", "0.00", ""),
          (None, "Agent A: share of pulls on arm 3 (best)", "=(«n3A»+IF(«armA»=3,1,0))/100", "0%", ""),
          (None, "Agent B: share of pulls on arm 3 (best)", "=(«n3B»+IF(«armB»=3,1,0))/100", "0%", "")])
    s.gap()
    rng_ = random.Random(2)
    us = [(round(rng_.random(), 3), round(rng_.random(), 3), round(rng_.random(), 3)) for _ in range(100)]
    hdr = ["Step t", "u1 (explore if < ε)", "u2 (random arm)", "u3 (win if < μ)"]
    for x in ("A", "B"):
        hdr += [f"{x}: arm", f"{x}: reward", f"{x}: Q1", f"{x}: Q2", f"{x}: Q3", f"{x}: N1", f"{x}: N2", f"{x}: N3",
                f"{x}: avg reward so far"]
    first = s.r + 1

    def row(i, r):
        out = [i + 1, *us[i]]
        for base, eps in ((4, "epsA"), (13, "epsB")):
            A_, R_ = col(base), col(base + 1)
            Qc = [col(base + 2 + k) for k in range(3)]
            Nc = [col(base + 5 + k) for k in range(3)]
            if i == 0:
                Q = ["=«q0»"] * 3
                N = ["=0"] * 3
            else:
                p = r - 1
                N = [f"={Nc[k]}{p}+IF({A_}{p}={k + 1},1,0)" for k in range(3)]
                Q = [f"={Qc[k]}{p}+IF({A_}{p}={k + 1},({R_}{p}-{Qc[k]}{p})/{Nc[k]}{r},0)" for k in range(3)]
            qr = f"{Qc[0]}{r}:{Qc[2]}{r}"
            out += [f"=IF(C{r}<«{eps}»,INT(D{r}*3)+1,MATCH(MAX({qr}),{qr},0))", f"=IF(E{r}<INDEX(«mu»,{A_}{r}),1,0)"]
            out += Q + N + [f"=AVERAGE({R_}${first}:{R_}{r})"]
        return out

    f1, f2 = s.table(hdr, 100, row, fmts={**{j: "0.00" for j in range(4, 22)}, 1: "0.000", 2: "0.000", 3: "0.000",
                                          0: "0", 4: "0", 5: "0", 9: "0", 10: "0", 11: "0", 13: "0", 14: "0",
                                          18: "0", 19: "0", 20: "0"},
                     label_cols=(0, 1, 2, 3), header_height=44)
    ch = line_chart("Average reward so far", "step", "average reward", width=15, height=8)
    for c in (FIRST + 12, FIRST + 21):
        ch.add_data(Reference(s.ws, min_col=c, min_row=f1 - 1, max_row=f2), titles_from_data=True)
    ch.set_categories(Reference(s.ws, min_col=FIRST, min_row=f1, max_row=f2))
    s.ws.add_chart(ch, f"Y{f1 - 1}")
    s.names.update(avgA=f"$N${f2}", avgB=f"$W${f2}", n3A=f"$M${f2}", n3B=f"$V${f2}", armA=f"$F${f2}",
                   armB=f"$O${f2}")
    s.note("What to notice: greedy agent A pulls arm 1 first, gets a few wins, and never tries anything else. Its "
           "average sits near 0.2. Agent B explores 10% of the time, finds arm 3 and averages about 0.8.")


def s07(wb):
    s = new(wb, "07 UCB", "Bandits & exploration", "Upper confidence bound (UCB)",
            "Explore smartly: add a bonus to arms you've tried less, so uncertainty itself attracts the agent. The bonus shrinks as evidence builds.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "ε-greedy explores blindly: it is as likely to re-try a clearly bad arm as a promising, barely-tested one.",
        "UCB scores each arm as estimate + uncertainty bonus, then picks the highest score. 'Optimism in the face "
        "of uncertainty': an arm might be great because we haven't tried it much.",
        "The bonus grows slowly with total time t (ln t) and shrinks as the arm is tried more (1/√N). Well-tested "
        "arms are judged on their estimate; rarely-tried arms get the benefit of the doubt.",
        "Untried arms (N = 0) are tried first. c sets how adventurous the agent is. (UCB1: Auer, Cesa-Bianchi & "
        "Fischer, 2002.)",
    ])
    s.section("MATH", "math")
    s.math(["UCB score(a) = Q(a) + c · √( ln t / N(a) )          choose the arm with the highest score"])
    s.legend([("t", "total pulls so far"), ("N(a)", "times arm a has been pulled"), ("c", "exploration strength")])
    s.section("EXAMPLE A: which arm next, after 20 pulls?", "example")
    s.kv([("t", "Total pulls so far t", 20, "0", ""), ("c", "Exploration strength c", 1.0, "0.00", "try 0.05")])
    data = [("Arm 1", 0.30, 10), ("Arm 2", 0.55, 8), ("Arm 3", 0.50, 2)]
    a, b = s.r + 1, s.r + 3
    s.table(["Arm", "Estimate Q", "Pulls N", "Bonus c·√(ln t / N)", "UCB score", "UCB choice", "Greedy choice"], 3,
            lambda i, r: list(data[i]) + [f"=«c»*SQRT(LN(«t»)/D{r})", f"=C{r}+E{r}",
                                          f'=IF(F{r}=MAX($F${a}:$F${b}),"◀ UCB picks this","")',
                                          f'=IF(C{r}=MAX($C${a}:$C${b}),"◀ greedy picks this","")'],
            fmts={1: "0.00", 2: "0", 3: "0.000", 4: "0.000"})
    s.note("What to notice: greedy picks arm 2 (best estimate). UCB picks arm 3: its estimate is only slightly lower "
           "but it has been tried just twice, so it might be better. UCB keeps preferring arm 3 until c drops below "
           "about 0.08; try c = 0.05 and it agrees with greedy.")
    s.section("EXAMPLE B: the bonus shrinks as an arm is tried more", "example")
    f1, f2 = s.table(["Pulls N", "Bonus at t above"], 20, lambda i, r: [i + 1, f"=«c»*SQRT(LN(«t»)/B{r})"],
                     fmts={0: "0", 1: "0.000"})
    ch = line_chart("Exploration bonus vs number of pulls", "pulls N", "bonus", width=13, height=8)
    ch.add_data(Reference(s.ws, min_col=FIRST + 1, min_row=f1 - 1, max_row=f2), titles_from_data=True)
    ch.set_categories(Reference(s.ws, min_col=FIRST, min_row=f1, max_row=f2))
    s.ws.add_chart(ch, f"E{f1 - 1}")


def s08(wb):
    s = new(wb, "08 Value Functions", "Planning with a model", "Value functions V(s) and Q(s, a)",
            "The value of a state is the return you expect from there onwards. Values let you compare states and actions without replaying the future each time.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "State value V_π(s) = expected return when starting in s and following policy π. It summarises the whole "
        "future in one number.",
        "Action value Q_π(s, a) = expected return when taking action a in s first, then following π. With Q you "
        "can choose actions directly: pick the a with the highest Q.",
        "Values depend on the policy: the same cell is worth a lot if you then head to the big exit, and little if "
        "you wander. The best possible values are written V* and Q*.",
        "Link between them: V_π(s) = Σ_a π(a | s) · Q_π(s, a).",
    ])
    s.section("MATH", "math")
    s.math([
        "V_π(s) = E_π[ G_t | s_t = s ]          Q_π(s, a) = E_π[ G_t | s_t = s, a_t = a ]",
        "For a deterministic corridor policy the expectation is just one path:  V(s) = r + γ · V(next cell)",
    ])
    s.section("EXAMPLE: values of 'always go right' vs 'always go left' in the corridor", "example")
    corridor_kv(s)
    a, b = s.r + 1, s.r + 5
    s.table(["Cell s", "V(s), always right", "V(s), always left", "Better policy from this cell"], 5,
            lambda i, r: [i + 1, "=«rR»" if i == 4 else f"=«g»*C{r + 1}", "=«rL»" if i == 0 else f"=«g»*D{r - 1}",
                          f'=IF(C{r}>D{r},"go right →",IF(C{r}<D{r},"← go left","tie"))'],
            fmts={0: "0", 1: "0.000", 2: "0.000"})
    s.gap()
    corridor_strip(s, ['="EXIT +"&«rL»'] + [f'="best V = "&TEXT(MAX(C{a + k},D{a + k}),"0.00")' for k in range(5)]
                   + ['="EXIT +"&«rR»'], label="Best value per cell", height=34)
    s.note("What to notice: at γ = 0.9 going right is better everywhere, even from cell 1 (10 × 0.9⁴ = 6.56 > 1). "
           "Set γ = 0.5 and cell 1 prefers the nearby +1 exit (1 > 0.625). Values turn a long-term question into a local "
           "comparison.")


def s09(wb):
    s = new(wb, "09 Policy Evaluation", "Planning with a model", "Bellman expectation equation and policy evaluation",
            "A state's value equals the expected immediate reward plus γ times the value of where you land next. Applying that rule repeatedly computes every value.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Bellman expectation equation: V_π(s) = average over actions and outcomes of [reward + γ · V_π(next "
        "state)]. Every value is defined in terms of neighbouring values.",
        "Iterative policy evaluation: start with V = 0 everywhere, then sweep through the states replacing each "
        "value with the right-hand side. Repeat until the values stop changing.",
        "This needs the model (P and R), so it is planning, not learning. Sheets 11–12 estimate the same values "
        "from experience instead.",
        "Policy here: in every cell, move right with probability p and left with probability 1 − p.",
    ])
    s.section("MATH", "math")
    s.math([
        "V_π(s) = Σ_a π(a | s) Σ_s' P(s' | s, a) · [ R + γ · V_π(s') ]",
        "Corridor version:  V_k+1(s) = p · [R(right) + γ·V_k(s+1)] + (1 − p) · [R(left) + γ·V_k(s−1)]      (V = 0 at exits)",
    ])
    s.section("EXAMPLE: evaluating the random policy, sweep by sweep", "example")
    corridor_kv(s)
    s.kv([("p", "P(move right) in every cell", 0.5, "0.00", "0.5 = random walk; try 0.9")])

    def row(i, r):
        if i == 0:
            return [0, "=0", "=0", "=0", "=0", "=0", ""]
        p = r - 1
        vals = []
        for st in range(1, 6):
            right = "«rR»" if st == 5 else f"«g»*{col(st + 1)}{p}"
            left = "«rL»" if st == 1 else f"«g»*{col(st - 1)}{p}"
            vals.append(f"=«p»*({right})+(1-«p»)*({left})")
        change = "=MAX(" + ",".join(f"ABS({col(k)}{r}-{col(k)}{p})" for k in range(1, 6)) + ")"
        return [i] + vals + [change]

    f1, f2 = s.table(["Sweep k", "V(1)", "V(2)", "V(3)", "V(4)", "V(5)", "Biggest change"], 51, row,
                     fmts={0: "0", **{j: "0.000" for j in range(1, 7)}})
    s.names["V3_final"] = f"$E${f2}"
    ch = line_chart("Values converge as sweeps repeat", "sweep k", "V(s)", width=15, height=8)
    ch.add_data(Reference(s.ws, min_col=FIRST + 1, max_col=FIRST + 5, min_row=f1 - 1, max_row=f2),
                titles_from_data=True)
    ch.set_categories(Reference(s.ws, min_col=FIRST, min_row=f1, max_row=f2))
    s.ws.add_chart(ch, f"J{f1 - 1}")
    s.note("What to notice: values spread outward from the exits sweep by sweep and settle within a few dozen sweeps "
           "(the biggest change shrinks toward 0). Even a random walker's cells nearer the +10 exit are worth more.")
    return f"'{s.name}'!$E${f2}"


def s10(wb):
    s = new(wb, "10 Value Iteration", "Planning with a model", "Bellman optimality and value iteration",
            "Replace 'average over the policy' with 'take the best action', repeat until values settle, and the best policy can be read straight off.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Bellman optimality equation: the best value of a state is the best action's [reward + γ · best value of "
        "the next state]. It defines the optimal values V* without fixing a policy first.",
        "Value iteration: start at V = 0 and repeatedly apply that 'max' update. Values converge to V*.",
        "The optimal policy is then greedy with respect to V*: in each state, take the action with the highest "
        "reward + γ · V*(next).",
        "Policy iteration is the alternative: evaluate a policy (sheet 09), improve it greedily, repeat. Both reach "
        "the same optimal policy.",
    ])
    s.section("MATH", "math")
    s.math([
        "V*(s) = max_a Σ_s' P(s' | s, a) · [ R + γ · V*(s') ]",
        "Value iteration:  V_k+1(s) = max_a [ R(s, a) + γ · V_k(s') ]          π*(s) = argmax_a [ R(s, a) + γ · V*(s') ]",
    ])
    s.section("EXAMPLE: optimal values and policy in the corridor", "example")
    corridor_kv(s)
    s.kv([("step", "Reward for every non-exit move", 0, "0.0", "try −2: a cost for wandering")])

    def row(i, r):
        if i == 0:
            return [0, "=0", "=0", "=0", "=0", "=0", ""]
        p = r - 1
        vals = []
        for st in range(1, 6):
            right = "«rR»" if st == 5 else f"«step»+«g»*{col(st + 1)}{p}"
            left = "«rL»" if st == 1 else f"«step»+«g»*{col(st - 1)}{p}"
            vals.append(f"=MAX({left},{right})")
        change = "=MAX(" + ",".join(f"ABS({col(k)}{r}-{col(k)}{p})" for k in range(1, 6)) + ")"
        return [i] + vals + [change]

    f1, f2 = s.table(["Sweep k", "V(1)", "V(2)", "V(3)", "V(4)", "V(5)", "Biggest change"], 13, row,
                     fmts={0: "0", **{j: "0.000" for j in range(1, 7)}})
    s.sub("Reading off the optimal policy from the final values")

    def qrow(i, r):
        st = i + 1
        left = "=«rL»" if st == 1 else f"=«step»+«g»*{col(st - 1)}{f2}"
        right = "=«rR»" if st == 5 else f"=«step»+«g»*{col(st + 1)}{f2}"
        return [st, left, right, f'=IF(D{r}>C{r},"→",IF(D{r}<C{r},"←","↔"))']

    qa, qb = s.r + 1, s.r + 5
    s.table(["Cell s", "Q(s, left)", "Q(s, right)", "Best action"], 5, qrow, fmts={0: "0", 1: "0.000", 2: "0.000"})
    s.gap()
    corridor_strip(s, ['="EXIT +"&«rL»'] + [f"=E{qa + k}" for k in range(5)] + ['="EXIT +"&«rR»'],
                   label="Optimal policy")
    s.note("What to notice: values stop changing after 5 sweeps (one per cell of distance) and the policy is → "
           "everywhere. Set γ = 0.5, or a step reward of −2, and cell 1 flips to ← (the +1 exit is one step away). "
           "With a step reward of −4, cells 1 and 2 both head for the nearby exit.")


def s11(wb, v3_ref):
    s = new(wb, "11 Monte Carlo", "Learning from experience", "Monte Carlo: learn values by averaging real returns",
            "No model needed: play whole episodes, record the return from each state, and average. The average converges to the true value.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Model-free learning: the agent doesn't know P or R. It only sees what actually happens when it acts.",
        "Monte Carlo (MC): run an episode to the end, compute the actual return G from each visited state, and "
        "average those returns. V(s) ≈ average of returns observed from s.",
        "Unbiased (it averages real returns) but noisy: one lucky or unlucky episode moves the estimate a lot, "
        "and you must wait until the episode ends before learning anything.",
        "Same update pattern as sheet 05: V ← V + (1/N)(G − V).",
    ])
    s.section("MATH", "math")
    s.math([
        "V(s) ≈ (1/N) · Σ_i G^(i)(s)          (average of N observed returns from state s)",
        "Incremental:  V(s) ← V(s) + (1/N) · (G − V(s))",
        "Corridor return from the start:  G = (exit reward) · γ^(T − 1)   for an episode of T steps",
    ])
    s.section("EXAMPLE: estimating V(3) for the random walk from 8 recorded episodes", "example")
    s.kv([("g", "Discount factor γ", 0.9, "0.00", "")])
    eps = random_walk(2, 8)
    data = [(len(e), e[-1][2]) for e in eps]
    f1, f2 = s.table(["Episode", "Steps T", "Exit reward", "Return G = reward × γ^(T−1)", "Running average V(3)",
             "Incremental V(3) (check)"], len(data),
            lambda i, r: [i + 1, data[i][0], data[i][1], f"=D{r}*«g»^(C{r}-1)", f"=AVERAGE($E${r - i}:E{r})",
                          f"=E{r}" if i == 0 else f"=G{r - 1}+(E{r}-G{r - 1})/B{r}"],
            fmts={0: "0", 1: "0", 2: "0", 3: "0.000", 4: "0.000", 5: "0.000"}, header_height=30)
    s.gap()
    s.kv([("v3", "True V(3) for this policy (sheet 09, p = 0.5)", f"={v3_ref}", "0.000",
           "planning answer, from the model"),
          (None, "MC estimate after 8 episodes", f"=F{f2}", "0.000", "learning answer, from experience")])
    s.note("Episodes were generated by a random walk (p = 0.5) with a fixed seed. What to notice: individual returns "
           "range from about 0.2 to 8.1. The average after 8 episodes is in the right area but noisy; MC needs many "
           "episodes. (If you change p on sheet 09 the 'true' value changes but these episodes don't.)")


def s12(wb, true_refs):
    s = new(wb, "12 TD Learning", "Learning from experience", "Temporal-difference learning, TD(0)",
            "Update after every single step, using your current guess of the next state's value as a stand-in for the rest of the future.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "TD learning doesn't wait for the episode to end. After each step it nudges V(s) toward r + γ·V(s'): the "
        "reward just received plus the current estimate of the next state.",
        "Using an estimate to update an estimate is called bootstrapping. It makes TD faster and less noisy than "
        "Monte Carlo, at the cost of some bias while the estimates are still wrong.",
        "TD error δ = r + γ·V(s') − V(s) = 'how much better or worse things went than expected'. It's the core "
        "signal of Q-learning, actor-critic (sheet 16) and much of neuroscience-inspired RL.",
        "Like MC, it is model-free: it only needs observed transitions (s, r, s').",
    ])
    s.section("MATH", "math")
    s.math([
        "δ_t = r_t+1 + γ · V(s_t+1) − V(s_t)          (V = 0 at exits)",
        "V(s_t) ← V(s_t) + α · δ_t",
    ])
    s.section("EXAMPLE: TD(0) on random-walk transitions, one row per step", "example")
    s.kv([("alpha", "Step size α", 0.3, "0.00", ""), ("g", "Discount factor γ", 0.9, "0.00", "")])
    trs = [(ei + 1, *t) for ei, e in enumerate(random_walk(103, 8)) for t in e]
    n = len(trs)
    first = s.r + 1

    def row(i, r):
        if i == n:
            vals = [f"={col(5 + k)}{r - 1}+IF(D{r - 1}={k + 1},«alpha»*L{r - 1},0)" for k in range(5)]
            return ["final", "", "", "", ""] + vals + [""]
        ep, st, _a, rew, s2 = trs[i]
        if i == 0:
            vals = ["=0"] * 5
        else:
            vals = [f"={col(5 + k)}{r - 1}+IF(D{r - 1}={k + 1},«alpha»*L{r - 1},0)" for k in range(5)]
        vr = f"G{r}:K{r}"
        delta = f"=E{r}+«g»*IF(OR(F{r}=0,F{r}=6),0,INDEX({vr},F{r}))-INDEX({vr},D{r})"
        return [i + 1, ep, st, rew, s2] + vals + [delta]

    f1, f2 = s.table(["Step", "Episode", "State s", "Reward r", "Next state s'", "V(1)", "V(2)", "V(3)", "V(4)",
                      "V(5)", "TD error δ"], n + 1, row,
                     fmts={0: "0", 1: "0", 2: "0", 3: "0", 4: "0", **{j: "0.000" for j in range(5, 11)}},
                     label_cols=(0, 1, 2, 3, 4), bold_rows=(n,))
    ch = line_chart("TD estimates over time", "step", "V(s)", width=15, height=8)
    ch.add_data(Reference(s.ws, min_col=FIRST + 5, max_col=FIRST + 9, min_row=f1 - 1, max_row=f2),
                titles_from_data=True)
    ch.set_categories(Reference(s.ws, min_col=FIRST, min_row=f1, max_row=f2))
    s.ws.add_chart(ch, f"N{f1 - 1}")
    s.sub("TD estimate vs the true values (sheet 09)")
    s.table(["Cell", "TD estimate", "True V (sheet 09)", "Error"], 5,
            lambda i, r: [i + 1, f"={col(5 + i)}{f2}", f"={true_refs[i]}", f"=C{r}-D{r}"],
            fmts={0: "0", 1: "0.000", 2: "0.000", 3: "0.000"})
    s.note(f"Transitions come from {8} random-walk episodes ({n} steps) with a fixed seed. What to notice: cells "
           "near the +10 exit learn first, because that's where the reward arrives. Values creep backwards one "
           "step at a time; more episodes (or a larger α) get closer to the true values.")


def s13(wb):
    s = new(wb, "13 Q-Learning", "Learning from experience", "Q-learning: learn the best action values from any behaviour",
            "Learn Q(s, a) with a TD update whose target assumes you'll act optimally next. It finds the optimal policy even while behaving randomly.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Q-learning (Watkins, 1989) learns action values Q(s, a) directly, so the policy is simply 'pick the "
        "action with the highest Q'. No model, no separate policy table.",
        "Its target uses the BEST next action: r + γ · max_a' Q(s', a'). So it learns the values of the optimal "
        "policy even when the data comes from a different, exploratory behaviour. That's called off-policy "
        "learning.",
        "Here the agent behaves completely at random, yet Q-learning recovers the optimal values Q* that "
        "value iteration computed with the model (sheet 10).",
        "Deep Q-Networks (DQN) are this same update with a neural network in place of the table.",
    ])
    s.section("MATH", "math")
    s.math([
        "target = r + γ · max_a' Q(s', a')          (max = 0 at exits)",
        "Q(s, a) ← Q(s, a) + α · (target − Q(s, a))",
    ])
    s.section("EXAMPLE: Q-learning from random behaviour in the corridor", "example")
    corridor_kv(s)
    s.kv([("alpha", "Step size α", 0.5, "0.00", "")])
    trs = [t for e in random_walk(203, 20) for t in e]
    n = len(trs)
    qcols = [col(5 + j) for j in range(10)]  # G..P: Q(1,L), Q(1,R), ..., Q(5,L), Q(5,R)
    qr = lambda r: f"G{r}:P{r}"  # noqa: E731

    def upd(j, r):
        st, act = j // 2 + 1, (-1 if j % 2 == 0 else 1)
        return f"={qcols[j]}{r - 1}+IF(AND(C{r - 1}={st},D{r - 1}={act}),«alpha»*R{r - 1},0)"

    def row(i, r):
        if i == n:
            return ["final", "", "", "", ""] + [upd(j, r) for j in range(10)] + ["", ""]
        st, act, rew, s2 = trs[i]
        qs = ["=0"] * 10 if i == 0 else [upd(j, r) for j in range(10)]
        target = (f"=E{r}+«g»*IF(OR(F{r}=0,F{r}=6),0,MAX(INDEX({qr(r)},(F{r}-1)*2+1),"
                  f"INDEX({qr(r)},(F{r}-1)*2+2)))")
        delta = f"=Q{r}-INDEX({qr(r)},(C{r}-1)*2+IF(D{r}=1,2,1))"
        return [i + 1, st, act, rew, s2] + qs + [target, delta]

    hdr = ["Step", "State s", "Action a", "Reward r", "Next s'"]
    for st in range(1, 6):
        hdr += [f"Q({st}, ←)", f"Q({st}, →)"]
    hdr += ["Target", "TD error"]
    f1, f2 = s.table(hdr, n + 1, row, fmts={0: "0", 1: "0", 2: "0", 3: "0", 4: "0",
                                             **{j: "0.00" for j in range(5, 17)}},
                     label_cols=(0, 1, 2, 3, 4), bold_rows=(n,))
    s.sub("Learned policy vs the optimal one")
    s.note("Optimal values Q* are computed directly from the corridor's rules (as value iteration on sheet 10 would), "
           "assuming non-exit moves pay 0.")
    vstar = lambda x: f"MAX(«rR»*«g»^(5-{x}),«rL»*«g»^({x}-1))"  # noqa: E731

    def prow(i, r):
        st = i + 1
        ql, qrr = f"={qcols[2 * i]}{f2}", f"={qcols[2 * i + 1]}{f2}"
        qsl = "=«rL»" if st == 1 else f"=«g»*{vstar(st - 1)}"
        qsr = "=«rR»" if st == 5 else f"=«g»*{vstar(st + 1)}"
        return [st, ql, qrr, f'=IF(D{r}>C{r},"→",IF(D{r}<C{r},"←","↔"))', qsl, qsr,
                f'=IF(G{r}>F{r},"→",IF(G{r}<F{r},"←","↔"))']

    s.table(["Cell", "Learned Q(s, ←)", "Learned Q(s, →)", "Learned best", "Optimal Q*(s, ←)", "Optimal Q*(s, →)",
             "Optimal best"], 5, prow, fmts={0: "0", 1: "0.00", 2: "0.00", 4: "0.00", 5: "0.00"}, header_height=30)
    s.note(f"Behaviour: {20} random-walk episodes ({n} transitions, fixed seed). What to notice: although the agent "
           "never tried to go right on purpose, the learned →-values approach 6.56, 7.29, 8.1, 9, 10, the optimal "
           "ones, and the learned policy is → in every cell.")


def s14(wb):
    s = new(wb, "14 SARSA vs Q-Learning", "Learning from experience", "On-policy vs off-policy: SARSA vs Q-learning",
            "Both update Q with a TD target. SARSA uses the action you actually take next; Q-learning uses the best one. That one difference changes what they learn.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "SARSA (state, action, reward, next state, next action) is on-policy: its target uses the next action the "
        "agent really takes, including exploratory ones. It learns the value of the behaviour it actually follows.",
        "Q-learning is off-policy: its target assumes the best next action, so it learns the optimal policy's "
        "values regardless of how the agent explores.",
        "Classic illustration, cliff walking: a path along a cliff edge is shortest but exploration occasionally "
        "steps off. Q-learning learns the edge path (optimal if you never slip); SARSA learns a safer path "
        "because it accounts for its own random slips.",
        "Expected SARSA averages over the next action's probabilities instead of using one sample, reducing noise.",
    ])
    s.section("MATH", "math")
    s.math([
        "SARSA:            target = r + γ · Q(s', a')          a' = the action actually chosen next",
        "Q-learning:       target = r + γ · max_a' Q(s', a')",
        "Expected SARSA:   target = r + γ · Σ_a' π(a' | s') · Q(s', a')",
        "All three:        Q(s, a) ← Q(s, a) + α · (target − Q(s, a))",
    ])
    s.section("EXAMPLE: the same transition updated three ways", "example")
    s.kv([("Q", "Current Q(s, a)", 2.0, "0.00", ""), ("r", "Reward received", 0, "0", ""),
          ("g", "Discount factor γ", 0.9, "0.00", ""), ("alpha", "Step size α", 0.5, "0.00", ""),
          ("QL", "Q(s', ←)", 1.0, "0.00", "next state's action values"), ("QR", "Q(s', →)", 8.0, "0.00", ""),
          ("anext", "Action actually chosen next (−1 = ←, +1 = →)", -1, "0", "an exploratory move"),
          ("eps", "ε of the ε-greedy behaviour", 0.2, "0.00", "")])
    exp_q = ("IF(«QR»>=«QL»,(1-«eps»+«eps»/2)*«QR»+«eps»/2*«QL»,(1-«eps»+«eps»/2)*«QL»+«eps»/2*«QR»)")
    rows = [("SARSA", "uses the action actually taken", "=«r»+«g»*IF(«anext»=1,«QR»,«QL»)"),
            ("Q-learning", "uses the best next action", "=«r»+«g»*MAX(«QL»,«QR»)"),
            ("Expected SARSA", "averages over ε-greedy choices", f"=«r»+«g»*{exp_q}")]
    s.table(["Method", "Target uses", "Target", "TD error", "New Q(s, a)"], 3,
            lambda i, r: list(rows[i]) + [f"=D{r}-«Q»", f"=«Q»+«alpha»*E{r}"],
            fmts={2: "0.000", 3: "0.000", 4: "0.000"}, label_cols=(0, 1), row_height=28)
    s.note("What to notice: the next move was an exploratory ←. SARSA lowers Q(s, a) (its real behaviour sometimes "
           "wanders off), Q-learning raises it a lot (it assumes you'll go → next time), and Expected SARSA lands in "
           "between.")


def s15(wb):
    s = new(wb, "15 Policy Gradient", "Policies & modern RL", "Policy gradient: learn the policy directly",
            "Instead of learning values and picking the max, adjust the policy's probabilities directly so rewarding actions become more likely.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Value-based methods (sheets 08–14) learn values and act greedily. Policy-gradient methods keep a "
        "parameterised stochastic policy and nudge its parameters uphill on expected reward.",
        "Why: they handle huge or continuous action spaces, learn genuinely random strategies when those are best, "
        "and are what's used to fine-tune language models (sheet 18).",
        "Softmax policy: each action has a preference θ; probabilities are e^θ normalised. The update raises θ for "
        "actions that did better than average and lowers it for the others.",
        "REINFORCE (Williams, 1992) estimates the gradient from samples: after taking action a and getting reward "
        "R, push θ in the direction (R − baseline) · (one-hot(a) − π).",
    ])
    s.section("MATH", "math")
    s.math([
        "π(a) = e^θ_a / Σ_b e^θ_b          J(θ) = Σ_a π(a) · q(a)   (expected reward)",
        "Exact gradient for a softmax:  ∂J/∂θ_a = π(a) · (q(a) − J)          update  θ_a ← θ_a + η · π(a) · (q(a) − J)",
        "Sampled (REINFORCE):  θ ← θ + η · (R − b) · (one-hot(a) − π)",
    ])
    s.section("EXAMPLE A: exact policy gradient on the 3-armed bandit", "example")
    s.kv([("eta", "Learning rate η", 2.0, "0.0", "")])
    a, b = s.r + 1, s.r + 3
    s.table(["Arm", "True value q(a)", "Starting preference θ0"], 3,
            lambda i, r: [f"Arm {i + 1}", MU[i], 0.0], fmts={1: "0.00", 2: "0.00"})
    s.gap()

    def row(i, r):
        th = [col(1 + k) for k in range(3)]
        S = "+".join(f"EXP({c}{r})" for c in th)
        if i == 0:
            thetas = [f"=$D${a + k}" for k in range(3)]
        else:
            p = r - 1
            Sp = "+".join(f"EXP({c}{p})" for c in th)
            thetas = [f"={th[k]}{p}+«eta»*EXP({th[k]}{p})/({Sp})*($C${a + k}-I{p})" for k in range(3)]
        probs = [f"=EXP({th[k]}{r})/({S})" for k in range(3)]
        return [i] + thetas + probs + [f"=F{r}*$C${a}+G{r}*$C${a + 1}+H{r}*$C${a + 2}"]

    f1, f2 = s.table(["Round", "θ1", "θ2", "θ3", "π(arm 1)", "π(arm 2)", "π(arm 3)", "Expected reward J"], 41, row,
                     fmts={0: "0", **{j: "0.000" for j in range(1, 8)}})
    ch = line_chart("Policy probabilities during training", "round", "probability", width=15, height=8)
    ch.add_data(Reference(s.ws, min_col=FIRST + 4, max_col=FIRST + 6, min_row=f1 - 1, max_row=f2),
                titles_from_data=True)
    ch.set_categories(Reference(s.ws, min_col=FIRST, min_row=f1, max_row=f2))
    s.ws.add_chart(ch, f"K{f1 - 1}")
    s.section("EXAMPLE B: one sampled REINFORCE update (what training actually sees)", "example")
    s.kv([("arm", "Arm that was pulled (1–3)", 3, "0", ""), ("R", "Reward received", 1, "0", ""),
          ("bl", "Baseline b (expected reward at round 0)", f"=I{f1}", "0.000", "sheet 16 explains baselines")])
    s.table(["Arm", "π at round 0", "One-hot of pulled arm", "Change in θ = η(R − b)(one-hot − π)"], 3,
            lambda i, r: [f"Arm {i + 1}", f"={col(4 + i)}{f1}", f"=IF({i + 1}=«arm»,1,0)",
                          f"=«eta»*(«R»-«bl»)*(D{r}-C{r})"],
            fmts={1: "0.000", 2: "0", 3: "0.000"}, header_height=30)
    s.note("What to notice: probability flows to arm 3, the best arm, and expected reward climbs toward 0.8. A single "
           "sampled win on arm 3 raises θ3 and lowers the others; a loss (set R = 0) does the opposite.")


def s16(wb):
    s = new(wb, "16 Actor-Critic", "Policies & modern RL", "Baselines, advantage and actor-critic",
            "Judge each action against what was expected: the advantage. A critic learns the expectation; the actor uses the advantage to update the policy.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Raw rewards make noisy policy updates: a mediocre action in a great state still gets a big reward. "
        "Subtracting a baseline (what you expected) fixes this without biasing the gradient.",
        "Advantage A(s, a) = Q(s, a) − V(s) = how much better this action was than the state's average.",
        "Actor-critic: the actor is the policy (sheet 15); the critic is a value function (sheet 12). The critic's "
        "TD error δ = r + γV(s') − V(s) is a one-step estimate of the advantage.",
        "This is the family behind A2C and PPO. GRPO, used for language models, removes the critic and uses the "
        "average reward of a group of answers as the baseline (sheet 18).",
    ])
    s.section("MATH", "math")
    s.math([
        "A(s, a) = Q(s, a) − V(s)   ≈   δ = r + γ · V(s') − V(s)",
        "Critic:  V(s) ← V(s) + α_critic · δ          Actor:  θ ← θ + η · δ · ∇ log π(a | s)",
    ])
    s.section("EXAMPLE: the critic scores five corridor moves", "example")
    corridor_kv(s)
    s.kv([("ac", "Critic step size α", 0.1, "0.00", "")])
    va, vb = s.r + 1, s.r + 5
    vals = [4.0, 5.0, 6.0, 7.5, 9.0]
    s.table(["Cell", "Critic's current V(s)"], 5, lambda i, r: [i + 1, vals[i]], fmts={0: "0", 1: "0.00"})
    s.names["V"] = f"$C${va}:$C${vb}"
    s.gap()
    moves = [(3, 1), (4, -1), (5, 1), (2, -1), (1, -1)]
    s.table(["State s", "Move", "Next s'", "Reward r", "V(s)", "V(s')", "TD error δ (advantage)", "Verdict",
             "Critic's new V(s)"], 5,
            lambda i, r: [moves[i][0], moves[i][1], f"=B{r}+C{r}", f"={reward_into(f'D{r}')}",
                          f"=INDEX(«V»,B{r})", f"=IF(OR(D{r}=0,D{r}=6),0,INDEX(«V»,D{r}))",
                          f"=E{r}+«g»*G{r}-F{r}",
                          f'=IF(H{r}>0,"better than expected → make it more likely","worse than expected → make it less likely")',
                          f"=F{r}+«ac»*H{r}"],
            fmts={0: "0", 1: "0", 2: "0", 3: "0", 4: "0.00", 5: "0.00", 6: "0.00", 8: "0.00"}, label_cols=(0,),
            row_height=28, header_height=30)
    s.note("What to notice: moving left from cell 1 earns +1, a positive reward, yet its advantage is −3 because the "
           "critic expected 4 from cell 1. The advantage, not the raw reward, decides the direction of the update.")


def s17(wb):
    s = new(wb, "17 Reward Design", "Policies & modern RL", "Reward design, shaping and reward hacking",
            "The agent optimises exactly the reward you write. Extra hints can speed learning, or create loopholes; potential-based shaping is the safe kind.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "The reward IS the task definition. If it can be satisfied in an unintended way, RL will find that way "
        "(reward hacking, a form of Goodhart's law).",
        "Sparse rewards (only at the goal) are honest but slow to learn from. Shaping adds intermediate rewards "
        "to guide the agent.",
        "Naive shaping is risky: a bonus for 'being near the goal' can teach the agent to hover near the goal "
        "forever instead of finishing.",
        "Potential-based shaping (Ng, Harada & Russell, 1999) gives F = γ·Φ(s') − Φ(s) for some 'progress' score Φ. "
        "Over any complete episode these bonuses telescope to −Φ(start), the same for every path, so the best "
        "policy cannot change and loops earn nothing.",
    ])
    s.section("MATH", "math")
    s.math([
        "Shaping bonus:  F(s, s') = γ · Φ(s') − Φ(s)          (Φ = 0 at exits)",
        "Shaped return = original return + Σ_t γ^t F_t = original return − Φ(s_0)",
        "Naive bonus b for arriving in cell 5:  hovering 5 → 4 → 5 → … is worth  b · γ / (1 − γ²)",
    ])
    s.section("EXAMPLE A: potential-based shaping keeps the ranking of paths", "example")
    corridor_kv(s)
    pa, pb = s.r + 1, s.r + 7
    phi = [0, 1, 2, 3, 4, 5, 0]
    s.table(["Cell", "Progress score Φ(s)"], 7, lambda i, r: [i, phi[i]], fmts={0: "0", 1: "0.0"})
    s.names["phi"] = f"$C${pa}:$C${pb}"
    s.gap()
    paths = {"A: straight to +10": [(3, 4), (4, 5), (5, 6)], "B: straight to +1": [(3, 2), (2, 1), (1, 0)],
             "C: wobble, then +10": [(3, 4), (4, 3), (3, 4), (4, 5), (5, 6)]}
    rows = [(name, t, st, s2) for name, steps in paths.items() for t, (st, s2) in enumerate(steps)]
    a, b = s.r + 1, s.r + len(rows)
    s.table(["Path", "Step t", "s", "s'", "Reward r", "γ^t", "γ^t · r", "Shaping F", "γ^t · (r + F)"], len(rows),
            lambda i, r: list(rows[i]) + [f"={reward_into(f'E{r}')}", f"=«g»^C{r}", f"=G{r}*F{r}",
                                          f"=«g»*INDEX(«phi»,E{r}+1)-INDEX(«phi»,D{r}+1)", f"=G{r}*(F{r}+I{r})"],
            fmts={1: "0", 2: "0", 3: "0", 4: "0", 5: "0.000", 6: "0.000", 7: "0.000", 8: "0.000"},
            label_cols=(0, 1, 2, 3))
    B, H, J = f"$B${a}:$B${b}", f"$H${a}:$H${b}", f"$J${a}:$J${b}"
    s.sub("Per path")
    s.table(["Path", "Original return", "Shaped return", "Shaped − original", "−Φ(start) (check)"], 3,
            lambda i, r: [list(paths)[i], f"=SUMIF({B},B{r},{H})", f"=SUMIF({B},B{r},{J})", f"=D{r}-C{r}",
                          "=-INDEX(«phi»,4)"],
            fmts={1: "0.000", 2: "0.000", 3: "0.000", 4: "0.000"}, label_cols=(0,))
    s.section("EXAMPLE B: a naive bonus that creates a loophole", "example")
    s.kv([("bonus", "Naive bonus for arriving in cell 5", 3, "0.0", "'reward being near the goal'"),
          ("hover", "Value of hovering 5 → 4 → 5 → … forever", "=«bonus»*«g»/(1-«g»^2)", "0.00", ""),
          ("finish", "Value of finishing from cell 5", "=«rR»", "0.00", "step into the +10 exit"),
          (None, "What the agent learns", '=IF(«hover»>«finish»,"HACK: hover forever","finish the task")', None, ""),
          (None, "Potential-based bonus for one 5 → 4 → 5 loop",
           "=(«g»*INDEX(«phi»,5)-INDEX(«phi»,6))+«g»*(«g»*INDEX(«phi»,6)-INDEX(«phi»,5))", "0.000",
           "negative: loops earn nothing extra")])
    s.note("What to notice: all three paths lose exactly Φ(3) = 3 under potential shaping, so A stays best. The naive "
           "+3 bonus makes hovering worth 14.2 > 10, so the agent never finishes. Lower the bonus to 2 and the "
           "loophole closes.")


def s18(wb):
    s = new(wb, "18 RL for LLMs", "Policies & modern RL", "From games to language models",
            "A language model is a policy over tokens. Reward a whole answer, compare answers to the same prompt, and push up the better ones. That's GRPO.")
    s.section("CONCEPT", "concept")
    s.bullets([
        "Mapping: state = the prompt plus the tokens written so far; action = the next token; episode = one full "
        "answer; policy = the language model's next-token probabilities.",
        "Reward usually arrives only at the end: a verifier (tests pass, maths answer correct), called RL with "
        "verifiable rewards, or a learned reward model trained from human preferences (RLHF).",
        "GRPO (Group Relative Policy Optimization): sample a group of answers to the same prompt; each answer's "
        "advantage is its reward minus the group average, divided by the group's spread. No critic needed: the "
        "group is the baseline (compare sheet 16).",
        "Practical additions: a KL penalty to stay close to the original model, clipping to limit each update's "
        "size, and care with reward hacking (sheet 17).",
        "For a deep dive on training coding agents this way, see the companion workbook RL_for_SWE_Agents.xlsx.",
    ])
    s.sub("Same ideas, two settings")
    pairs = [("State", "Cell in the corridor", "Prompt + tokens so far"),
             ("Action", "Move ← or →", "Next token"),
             ("Episode", "Start → exit", "Prompt → full answer"),
             ("Reward", "+1 / +10 at an exit", "1 if the answer is right (or a reward-model score)"),
             ("Policy", "Probabilities of ← / →", "The language model itself"),
             ("Baseline / advantage", "Critic V(s) (sheet 16)", "Group average reward (GRPO)")]
    s.table(["Concept", "Corridor world", "Language model"], len(pairs), lambda i, r: list(pairs[i]),
            label_cols=(0, 1, 2), merge_last=True, row_height=20)
    s.section("MATH", "math")
    s.math([
        "J(θ) = E_x E_{y ~ π_θ(·|x)} [ R(x, y) ]          (expected reward of answers y to prompts x)",
        "GRPO advantage:  A_i = (R_i − mean(R)) / std(R)   over a group of G answers to the same prompt",
        "Loss ≈ − mean_i [ A_i · (1/|y_i|) · log π_θ(y_i | x) ] + β · KL(π_θ ‖ π_ref)",
    ])
    s.section("EXAMPLE: one GRPO group for the prompt '2 + 3 × 4 = ?'", "example")
    s.kv([("ans", "Correct answer", 14, "0", "the verifier checks this")])
    answers = [14, 20, 14, 24]
    a, b = s.r + 1, s.r + 4
    s.table(["Sample", "Model's answer", "Reward (1 = correct)", "Advantage A", "Training effect"], 4,
            lambda i, r: [f"Answer {i + 1}", answers[i], f"=IF(C{r}=«ans»,1,0)",
                          f"=IF(«std»=0,0,(D{r}-«mean»)/«std»)",
                          f'=IF(E{r}>0,"make this answer more likely",IF(E{r}<0,"make this answer less likely","no change"))'],
            fmts={1: "0", 2: "0", 3: "0.00"}, label_cols=(0,))
    s.gap()
    s.kv([("mean", "Group mean reward", f"=AVERAGE(D{a}:D{b})", "0.00", "the baseline"),
          ("std", "Group std", f"=STDEVP(D{a}:D{b})", "0.00", "if 0 (all right or all wrong): nothing to learn")])
    s.note("What to notice: '20' is what you get by adding before multiplying, a common mistake. The two correct "
           "answers get +1 and the two wrong ones −1. Make all four answers 14 and every advantage becomes 0: GRPO "
           "only learns when answers differ.")


GLOSSARY = [
    ("Agent", "The learner that chooses actions.", "01 The RL Loop"),
    ("Environment", "Everything the agent interacts with; returns next states and rewards.", "01 The RL Loop"),
    ("State (s)", "The situation the agent is in.", "01 The RL Loop"),
    ("Action (a)", "A choice the agent can make.", "01 The RL Loop"),
    ("Reward (r)", "Number returned after an action; higher is better.", "01 The RL Loop"),
    ("Episode", "One run from a start to an end.", "01 The RL Loop"),
    ("Return (G)", "Total (discounted) reward from now to the end of the episode.", "02 Return & Discount"),
    ("Discount factor (γ)", "How much a reward one step later is worth (0–1).", "02 Return & Discount"),
    ("Markov property", "The future depends only on the current state and action.", "03 MDP"),
    ("MDP", "States, actions, transition probabilities, rewards and γ.", "03 MDP"),
    ("Transition probability", "P(s' | s, a): chance of landing in s' after a in s.", "03 MDP"),
    ("Policy (π)", "The agent's rule for choosing actions in each state.", "04 Policy"),
    ("Multi-armed bandit", "One state, several choices with unknown average payouts.", "05 Action Values"),
    ("Action value Q(a)", "Average reward of an action (bandit) — estimated from samples.", "05 Action Values"),
    ("Step size (α)", "How far each update moves an estimate toward its target.", "05 Action Values"),
    ("Exploration vs exploitation", "Trying new things to learn vs using what looks best now.", "06 Epsilon-Greedy"),
    ("ε-greedy", "Act randomly with probability ε, greedily otherwise.", "06 Epsilon-Greedy"),
    ("Optimistic initial values", "Start estimates high so everything gets tried early.", "06 Epsilon-Greedy"),
    ("UCB", "Pick the action with the best estimate + uncertainty bonus.", "07 UCB"),
    ("State value V(s)", "Expected return from state s under a policy.", "08 Value Functions"),
    ("Action value Q(s, a)", "Expected return after taking a in s, then following the policy.", "08 Value Functions"),
    ("Bellman expectation equation", "V(s) = expected [r + γ·V(s')] under the policy.", "09 Policy Evaluation"),
    ("Policy evaluation", "Computing V for a given policy by repeated Bellman updates.", "09 Policy Evaluation"),
    ("Bellman optimality equation", "V*(s) = max over actions of [r + γ·V*(s')].", "10 Value Iteration"),
    ("Value iteration", "Repeated max-Bellman updates; converges to optimal values.", "10 Value Iteration"),
    ("Policy iteration", "Alternate policy evaluation and greedy improvement.", "10 Value Iteration"),
    ("Model-based vs model-free", "Planning with known P and R vs learning from experience.", "03 MDP"),
    ("Monte Carlo", "Estimate values by averaging returns of complete episodes.", "11 Monte Carlo"),
    ("Bootstrapping", "Updating an estimate using another estimate.", "12 TD Learning"),
    ("TD error (δ)", "r + γ·V(s') − V(s): surprise relative to expectation.", "12 TD Learning"),
    ("Q-learning", "Off-policy TD control with target r + γ·max Q(s', ·).", "13 Q-Learning"),
    ("On-policy / off-policy", "Learning about the behaviour you follow vs a different (e.g. optimal) one.",
     "14 SARSA vs Q-Learning"),
    ("SARSA", "On-policy TD control using the next action actually taken.", "14 SARSA vs Q-Learning"),
    ("Policy gradient", "Adjust policy parameters directly to increase expected reward.", "15 Policy Gradient"),
    ("Softmax policy", "Probabilities proportional to e^θ for each action.", "15 Policy Gradient"),
    ("REINFORCE", "Sample-based policy gradient: (R − b)·∇log π.", "15 Policy Gradient"),
    ("Baseline", "Expected reward subtracted to reduce noise in updates.", "16 Actor-Critic"),
    ("Advantage", "How much better an action was than expected: Q − V.", "16 Actor-Critic"),
    ("Actor-critic", "Policy (actor) updated using a learned value function (critic).", "16 Actor-Critic"),
    ("Reward shaping", "Adding intermediate rewards to guide learning.", "17 Reward Design"),
    ("Potential-based shaping", "Shaping F = γΦ(s') − Φ(s) that never changes the best policy.", "17 Reward Design"),
    ("Reward hacking", "Exploiting a flaw in the reward instead of doing the task.", "17 Reward Design"),
    ("RLHF", "RL from human feedback via a learned reward model.", "18 RL for LLMs"),
    ("Verifiable reward", "Reward computed by an automatic checker (tests, exact answer).", "18 RL for LLMs"),
    ("GRPO", "Policy gradient with group-average baseline; no critic.", "18 RL for LLMs"),
    ("KL penalty", "Keeps the trained model close to its starting point.", "18 RL for LLMs"),
]


def glossary(wb):
    ws = wb.create_sheet("Glossary")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 70
    ws.column_dimensions["D"].width = 26
    link = ws.cell(1, 2, "← Back to Index")
    link.hyperlink = "#'Index'!A1"
    link.font = font(color="0563C1", underline="single", size=9)
    ws.merge_cells("B2:D2")
    t = ws["B2"]
    t.value = "Glossary"
    t.font = font(bold=True, size=16, color="FFFFFF")
    t.fill = fill(NAVY)
    ws.row_dimensions[2].height = 30
    for j, h in enumerate(["Term", "Plain meaning", "Explained on"]):
        c = ws.cell(4, 2 + j, h)
        c.font = font(bold=True)
        c.fill = fill(FILL["hdr"])
        c.border = BOX
    for i, (term, meaning, sheet) in enumerate(GLOSSARY):
        r = 5 + i
        ws.cell(r, 2, term).font = font(bold=True)
        ws.cell(r, 3, meaning).font = font()
        ws.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="center")
        lc = ws.cell(r, 4, sheet)
        lc.hyperlink = f"#'{sheet}'!A1"
        lc.font = font(color="0563C1", underline="single")
        for cc in range(2, 5):
            ws.cell(r, cc).border = BOX
    return ws


def index(wb):
    ws = wb["Index"]
    ws.sheet_view.showGridLines = False
    for letter, w in (("A", 2), ("B", 6), ("C", 34), ("D", 24), ("E", 92)):
        ws.column_dimensions[letter].width = w
    ws.merge_cells("B2:E2")
    c = ws["B2"]
    c.value = "Reinforcement Learning Fundamentals"
    c.font = font(bold=True, size=18, color="FFFFFF")
    c.fill = fill(NAVY)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[2].height = 34
    ws.merge_cells("B3:E3")
    ws["B3"].value = ("RL from the ground up with two tiny running examples. Every sheet goes CONCEPT (plain words) → "
                      "MATH (the formulas) → EXAMPLE (live Excel you can edit).")
    ws["B3"].font = font(italic=True, size=11, color=NAVY)
    ws["B3"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[3].height = 32
    r = 5
    ws.cell(r, 2, "Running examples").font = font(bold=True, size=11, color=NAVY)
    r += 1
    for text in ["Corridor world: cells 0–6, start in cell 3; cell 0 is a small exit (+1), cell 6 a big exit (+10). "
                 "Used for values, planning, TD, Q-learning, actor-critic and shaping.",
                 "3-armed bandit: three slot machines with hidden win rates 0.2, 0.5 and 0.8. Used for action values, "
                 "exploration and policy gradient."]:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        ws.cell(r, 2, "•  " + text).font = font()
        r += 1
    r += 1
    ws.cell(r, 2, "How to read the sheets").font = font(bold=True, size=11, color=NAVY)
    r += 1
    for sample, text, is_input in [("0.9", "Blue text on yellow = an input you can change. Everything else recalculates.", True),
                                   ("6.561", "Black text = a formula. Don't type over these.", False),
                                   ("", "All numbers are small illustrative examples chosen to make each idea visible. "
                                        "Simulated data uses fixed random seeds so results are reproducible.", False)]:
        cc = ws.cell(r, 2, sample)
        cc.alignment = Alignment(horizontal="center")
        cc.font = font(color="0000FF" if is_input else "000000")
        if sample:
            cc.border = BOX
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
    for i, (name, part, title, idea) in enumerate(SHEETS, 1):
        ws.cell(r, 2, i).font = font()
        ws.cell(r, 2).alignment = Alignment(horizontal="center", vertical="center")
        lc = ws.cell(r, 3, title)
        lc.hyperlink = f"#'{name}'!A1"
        lc.font = font(color="0563C1", underline="single")
        lc.alignment = Alignment(wrap_text=True, vertical="center")
        pc = ws.cell(r, 4, part)
        pc.font = font()
        pc.fill = fill(PARTS[part][0])
        pc.alignment = Alignment(vertical="center")
        ic = ws.cell(r, 5, idea)
        ic.font = font()
        ic.alignment = Alignment(wrap_text=True, vertical="center")
        for cc in range(2, 6):
            ws.cell(r, cc).border = BOX
        ws.row_dimensions[r].height = 30 if len(idea) > 100 else 18
        r += 1
    gl = ws.cell(r, 3, "Glossary of terms")
    gl.hyperlink = "#'Glossary'!A1"
    gl.font = font(color="0563C1", underline="single", bold=True)
    r += 2
    ws.cell(r, 2, "Suggested path").font = font(bold=True, size=11, color=NAVY)
    r += 1
    for t in ["1. Foundations (01–04): the loop, returns, MDPs and policies.",
              "2. Bandits & exploration (05–07): estimating values and the explore/exploit dilemma.",
              "3. Planning with a model (08–10): value functions and the Bellman equations.",
              "4. Learning from experience (11–14): Monte Carlo, TD, Q-learning, SARSA.",
              "5. Policies & modern RL (15–18): policy gradient, actor-critic, reward design, RL for language models."]:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        ws.cell(r, 2, t).font = font()
        r += 1


def main():
    wb = Workbook()
    wb.active.title = "Index"
    for f in (s01, s02, s03, s04, s05, s06, s07, s08):
        f(wb)
    v3_ref = s09(wb)
    s10(wb)
    s11(wb, v3_ref)
    pe = "'09 Policy Evaluation'"
    last = int(v3_ref.split("$")[-1])
    s12(wb, [f"{pe}!${col(k)}${last}" for k in range(1, 6)])
    for f in (s13, s14, s15, s16, s17, s18):
        f(wb)
    for sheet in SHEET_OBJS:
        sheet.resolve()
    glossary(wb)
    index(wb)
    for ws in wb.worksheets:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_setup.orientation = "landscape"
    sheets = {ws.title: ws for ws in wb.worksheets}
    for name, part, _, _ in SHEETS:
        sheets[name].sheet_properties.tabColor = PARTS[part][1]
    wb.save(OUT)
    print("saved", OUT)


if __name__ == "__main__":
    main()
