# RL for SWE Agents: Concept Workbook

`RL_for_SWE_Agents.xlsx` breaks down the ideas behind
[Mini-SWE-RL](https://github.com/RajatDandekar/Mini-SWE-RL) and its `swe-rl-ipr` follow-up.
There is one sheet per topic, and each sheet has three sections:

- **Concept**: the idea in plain words, plus where it lives in the source repo
- **Math**: the formulas and a symbol legend
- **Example**: a worked example built from live Excel formulas. Blue cells on yellow are inputs you can change.

The first sheet is an index. The 21 topics are grouped as Foundations (RL setup, reward, temperature,
log-probs, policy gradient, baselines), GRPO mechanics (advantage, difficulty sweet spot, loss),
Production recipe (KL, PPO clipping, leave-one-out, length normalisation, on/off-policy, multi-turn
agents) and Evaluation & reward integrity (statistical noise, reward hacking, IPR, precision/recall,
a raw-vs-IPR training simulation, compute and scale).

Rebuild the workbook with `python build_workbook.py` (requires `openpyxl`).
