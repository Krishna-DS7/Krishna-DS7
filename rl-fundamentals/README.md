# Reinforcement Learning Fundamentals Workbook

`RL_Fundamentals.xlsx` teaches reinforcement learning from scratch, one concept per sheet. Each sheet has:

- **Concept**: the idea in plain words
- **Math**: the formulas and a symbol legend
- **Example**: a worked example built from live Excel formulas. Blue cells on yellow are inputs you can change.

Two tiny running examples are used throughout:

- **Corridor world**: cells 0–6. The agent starts in cell 3; cell 0 is a small exit (+1) and cell 6 a big exit (+10).
- **3-armed bandit**: slot machines with hidden win rates 0.2, 0.5 and 0.8.

| Part | Sheets |
|---|---|
| Foundations | 01 RL loop · 02 Return & discount · 03 MDP · 04 Policy |
| Bandits & exploration | 05 Action values · 06 ε-greedy (100-step simulation) · 07 UCB |
| Planning with a model | 08 Value functions · 09 Policy evaluation · 10 Value iteration |
| Learning from experience | 11 Monte Carlo · 12 TD learning · 13 Q-learning · 14 SARSA vs Q-learning |
| Policies & modern RL | 15 Policy gradient · 16 Actor-critic · 17 Reward design · 18 RL for LLMs |

There is also a Glossary sheet. Simulated data (bandit draws, random-walk episodes) is generated with fixed seeds,
so the workbook is reproducible.

Rebuild with `python build_rl_fundamentals.py` (requires `openpyxl`). The script reuses the layout engine in
`../rl-swe-agents/build_workbook.py`. The companion workbook `../rl-swe-agents/RL_for_SWE_Agents.xlsx` covers RL
for coding agents in depth.
