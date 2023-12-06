# BejeweledBot

## Overview
BejeweledBot is an automated bot designed to play the Bejeweled game. The bot takes screenshots of the game grid, analyzes the colors in each cell, and strategically makes moves to achieve optimal matches.

Play the game here: [Bejeweled Classic](https://www.digbejeweled.com/bejeweled-classic.php)

## How it Works
The bot follows these steps to play the game:
1. **Capture Grid**: Takes screenshots of the game grid.
2. **Color Recognition**: Identifies colors in each cell using predefined color mappings.
3. **Move Optimization**: Searches for optimal moves based on the current color grid state using a heuristic-based approach.
4. **Move Execution**: Performs the identified optimal move by simulating mouse clicks.

## Installation
Before running the bot, make sure to install the required dependencies. You can do this by running the following command in your virtual environment:


```bash
pip install -r requirements.txt
```

The dependencies include:
- numpy: Numerical operations for array manipulation.
- opencv-python: Computer vision library for image processing.
- pyautogui: Automation library for simulating mouse and keyboard actions.
- keyboard: Library for working with keyboard inputs.

## Usage

- Setup Grid Coordinates: Run the script and follow the prompts to click on the top-left and bottom-right corners of the Bejeweled game grid.
- Game Automation: The bot will continuously capture the grid, analyze colors, and make optimal moves until you press and hold the "Escape" key.

Feel free to customize the color mappings and tweak the bot's logic based on your game environment.