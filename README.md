# BejeweledBot

## Overview
BejeweledBot is an automated bot that plays Bejeweled 3 (Steam). It captures the game grid via screenshots, identifies gem colors and special gems using HSV analysis, evaluates moves with a 2-step look-ahead and cascade simulation, and executes moves via mouse automation.

## How it Works
1. **Grid Detection**: Finds the Bejeweled 3 window and calculates the grid position using calibrated percentages (or manual corner clicks as fallback).
2. **Color Recognition**: Identifies gem colors using HSV color space analysis. Detects special gems (Flame, Star, Hypercube) by analyzing the border glow around each cell.
3. **Move Evaluation**: Scores every possible swap using Bejeweled 3's actual point values (match-3: 50, match-4: 100, match-5: 500, Star Gem: +150, cascades: stacking +50 bonus). A 2-step look-ahead evaluates what follow-up moves become available after cascades settle.
4. **Move Execution**: Performs the highest-scoring valid move via mouse clicks, waits for animations to finish, then repeats.

## Features
- **Special gem detection**: Identifies Flame Gems, Star Gems, and Hypercubes by their visual glow patterns
- **Cascade simulation**: Simulates gravity and chain reactions to find moves that trigger cascades
- **Board stability detection**: Waits for animations to finish before scanning
- **Game over detection**: Pauses when the game ends (press Space to resume, Escape to quit)
- **Stuck loop prevention**: Blacklists moves that repeatedly fail and tries different board areas
- **Playthrough logging**: Each session creates a timestamped log file in `logs/`
- **Gem library**: Automatically captures screenshots of gems for identification and review

## Installation

```bash
pip install -r requirements.txt
```

Dependencies: numpy, opencv-python, pyautogui, keyboard, pygetwindow

## Setup

### First run (calibration)
Run the calibration script once to save your grid position:

```bash
python calibrate.py
```

This finds the Bejeweled 3 window, asks you to click the top-left and bottom-right corners of the game grid, and saves the position as percentages to `grid_config.json`. The bot will use these automatically on every subsequent run.

If you skip calibration, the bot uses default percentages that work for most setups.

### Running the bot

```bash
python bejeweled.py
```

Or double-click `run.bat`.

### Controls
| Key | Action |
|---|---|
| **Escape** | Stop the bot |
| **Space** | Resume after game over |
| **P** | Snapshot all 64 cells for gem review |

## Gem Library

The bot automatically captures screenshots of gems it encounters in `gem_library/`. To review and classify special gems:

```bash
python review_gems.py
```

Keys during review:
- **Enter** = correct, keep as-is
- **r/o/y/g/b/p/w** = classify as color (regular gem)
- Color key + **f/s/h/n** = special type (e.g., `r` then `f` = red flame)
- **d** = delete, **q** = quit

## Logging

Each playthrough creates a log file in `logs/` with:
- Full grid state each frame (DEBUG level)
- Every move with score (INFO level)
- Special gem detections, blacklist events, board stability info

## Project Structure

```
bejeweled.py      # Main bot
calibrate.py      # One-time grid calibration
review_gems.py    # Gem screenshot reviewer
run.bat           # Windows launcher
grid_config.json  # Saved grid position (created by calibrate.py)
logs/             # Playthrough logs
gem_library/      # Captured gem screenshots
```

## License
This project is licensed under the GNU General Public License v3.0.
