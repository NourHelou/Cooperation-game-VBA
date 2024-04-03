
# Game Theory: Cooperation

## Overview

This project explores the concept of cooperation within the framework of game theory, drawing inspiration from both biological evolution and social sciences. It investigates the paradox of cooperation—why individuals in a competitive environment might choose to cooperate rather than act selfishly, despite natural selection traditionally favoring selfish behaviors.

## Objectives

- To understand the emergence and sustainability of cooperative behavior in the natural world and in human societies.
- To analyze conflict between collective interest and individual interests using mathematical models within the context of game theory.

## The Model: Prisoner's Dilemma

We utilize a simple game theory model known as the "Prisoner's Dilemma" where two players decide independently to either cooperate or defect without knowledge of the other's decision. The outcome rewards or penalizes them based on the combination of their choices, encouraging analysis of strategic decision-making in a controlled setting.

## Implementation

### Simulation Setup

- **Environment:** A 20x20 grid where 400 autonomous agents (robots) are placed.
- **Agents:** Each robot adheres to specific behavioral rules, categorized by complexity from simple (e.g., always cooperate or defect) to advanced strategies involving memory and adaptability.
- **Interactions:** Robots randomly interact with neighbors, with outcomes affecting their strategies and scores based on the rules of the Prisoner's Dilemma.

### Rules

We explored various strategies, including:
- **Basic Strategies:** Selfish, Generous.
- **Intermediate Strategies:** Familial, Moody, Sectarian.
- **Advanced Strategies:** Psychotic, Reputation-based.
- **Complex Strategies:** Tit-for-Tat, Forgiving, Elephant (memory-based), and Random Elephant.

### Simulation Phases

1. Initial setup with basic and easy rule-following robots.
2. Integration of intermediate complexity robots.
3. Implementation of a rule-change mechanism after 40 iterations based on performance.
4. Progressive integration of complex strategy robots and analysis of different distributions and probabilities in initial setup.
5. Optional adjustments to explore memory-based strategies' effectiveness by altering the simulation's scale and duration.

## Results and Analysis

- The simulation aims to identify which strategies yield the best and worst scores over time.
- Observations on the dynamics of rule changes and strategy effectiveness in various setups.
- Analysis of different initial distributions and their impact on the evolution of cooperative behaviors.
