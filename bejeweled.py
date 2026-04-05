"""
BejeweledBot - An automated bot that plays Bejeweled Classic.

Captures the game grid via screenshots, identifies gem colors, evaluates possible
moves using a heuristic scoring system, and executes the best move via mouse automation.
"""

import concurrent.futures
import logging
import os
import time
from datetime import datetime

import cv2
import keyboard
import numpy as np
import pyautogui

# Color mappings (BGR format) to gem names
COLORS = {
    (206, 165, 33): "blue",
    (41, 115, 16): "green",
    (16, 49, 239): "red",
    (206, 70, 196): "purple",
    (24, 123, 255): "orange",
    (24, 222, 255): "yellow",
    (211, 211, 211): "white",
}

COLOR_TOLERANCE = 3
GRID_SIZE = 8

# Single-letter abbreviations for log output
COLOR_ABBREV = {
    "blue": "B",
    "green": "G",
    "red": "R",
    "purple": "P",
    "orange": "O",
    "yellow": "Y",
    "white": "W",
}


def setup_logger():
    """Set up logging to both console and a timestamped log file."""
    os.makedirs("logs", exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_file = os.path.join("logs", f"playthrough_{timestamp}.log")

    logger = logging.getLogger("bejeweled")
    logger.setLevel(logging.DEBUG)

    file_handler = logging.FileHandler(log_file)
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    )

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(logging.Formatter("%(message)s"))

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    logger.info("Playthrough log: %s", log_file)
    return logger


def format_grid(color_grid):
    """Format the color grid as a compact string for logging."""
    lines = ["  " + " ".join(str(c) for c in range(GRID_SIZE))]
    for row_idx, row in enumerate(color_grid):
        cells = " ".join(COLOR_ABBREV.get(c, ".") for c in row)
        lines.append(f"{row_idx} {cells}")
    return "\n".join(lines)


def get_grid_coordinates():
    """Prompt the user to click the corners of the game grid."""
    input("Move your mouse to the top-left corner of the game grid and press Enter.")
    top_left = pyautogui.position()
    input("Move your mouse to the bottom-right corner of the game grid and press Enter.")
    bottom_right = pyautogui.position()
    return top_left, bottom_right


def capture_grid(top_left, bottom_right):
    """Capture a screenshot of the game grid and draw overlay grid lines."""
    screenshot = pyautogui.screenshot(
        region=(
            top_left[0],
            top_left[1],
            bottom_right[0] - top_left[0],
            bottom_right[1] - top_left[1],
        )
    )
    grid_img = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

    cell_width = grid_img.shape[1] // GRID_SIZE
    cell_height = grid_img.shape[0] // GRID_SIZE
    for i in range(1, GRID_SIZE):
        cv2.line(
            grid_img,
            (i * cell_width, 0),
            (i * cell_width, grid_img.shape[0]),
            (0, 255, 0),
            2,
        )
        cv2.line(
            grid_img,
            (0, i * cell_height),
            (grid_img.shape[1], i * cell_height),
            (0, 255, 0),
            2,
        )

    return grid_img


def identify_cell_color(grid_img, row, col):
    """Identify the gem color at the given grid cell. Returns (row, col, color_name)."""
    cell_width = grid_img.shape[1] // GRID_SIZE
    cell_height = grid_img.shape[0] // GRID_SIZE
    cell = grid_img[
        row * cell_height : (row + 1) * cell_height,
        col * cell_width : (col + 1) * cell_width,
    ]

    for color_bgr, name in COLORS.items():
        mask = cv2.inRange(
            cell,
            np.array(color_bgr) - COLOR_TOLERANCE,
            np.array(color_bgr) + COLOR_TOLERANCE,
        )
        if cv2.countNonZero(mask) > 0:
            return row, col, name

    return row, col, ""


def build_color_grid(grid_img):
    """Identify colors for all cells in parallel. Returns an 8x8 color grid."""
    color_grid = [["" for _ in range(GRID_SIZE)] for _ in range(GRID_SIZE)]

    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = [
            executor.submit(identify_cell_color, grid_img, row, col)
            for row in range(GRID_SIZE)
            for col in range(GRID_SIZE)
        ]
        for future in concurrent.futures.as_completed(futures):
            row, col, color = future.result()
            color_grid[row][col] = color

    return color_grid


def evaluate_state(grid):
    """Score the grid based on consecutive matching gems (3+ in a row/column)."""
    score = 0

    # Check rows
    for row in range(GRID_SIZE):
        col = 0
        while col < GRID_SIZE:
            color = grid[row][col]
            if not color:
                col += 1
                continue
            run_length = 1
            while col + run_length < GRID_SIZE and grid[row][col + run_length] == color:
                run_length += 1
            if run_length >= 3:
                score += run_length
            col += run_length

    # Check columns
    for col in range(GRID_SIZE):
        row = 0
        while row < GRID_SIZE:
            color = grid[row][col]
            if not color:
                row += 1
                continue
            run_length = 1
            while row + run_length < GRID_SIZE and grid[row + run_length][col] == color:
                run_length += 1
            if run_length >= 3:
                score += run_length
            row += run_length

    return score


def evaluate_move(color_grid, row, col, direction):
    """Evaluate a move on a copy of the grid. Returns (score, move_tuple)."""
    grid = [r[:] for r in color_grid]

    if direction == "right":
        grid[row][col], grid[row][col + 1] = grid[row][col + 1], grid[row][col]
        return evaluate_state(grid), (row, col, row, col + 1)

    grid[row][col], grid[row + 1][col] = grid[row + 1][col], grid[row][col]
    return evaluate_state(grid), (row, col, row + 1, col)


def find_optimal_move(color_grid):
    """Find the move producing the highest score. Returns (move_tuple, score) or (None, 0)."""
    best_move = None
    best_score = 0

    moves = []
    for row in range(GRID_SIZE - 1, -1, -1):
        for col in range(GRID_SIZE):
            if not color_grid[row][col]:
                continue
            if col < GRID_SIZE - 1:
                moves.append((color_grid, row, col, "right"))
            if row < GRID_SIZE - 1:
                moves.append((color_grid, row, col, "down"))

    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = [executor.submit(evaluate_move, *args) for args in moves]
        for future in concurrent.futures.as_completed(futures):
            score, move = future.result()
            if score > best_score:
                best_score = score
                best_move = move

    return best_move, best_score


def perform_move(top_left, grid_img, src_row, src_col, dest_row, dest_col):
    """Click the source and destination cells to perform a gem swap."""
    cell_width = grid_img.shape[1] // GRID_SIZE
    cell_height = grid_img.shape[0] // GRID_SIZE
    half_w = cell_width // 2
    half_h = cell_height // 2

    from_pos = (
        top_left[0] + src_col * cell_width + half_w,
        top_left[1] + src_row * cell_height + half_h,
    )
    to_pos = (
        top_left[0] + dest_col * cell_width + half_w,
        top_left[1] + dest_row * cell_height + half_h,
    )

    pyautogui.click(from_pos)
    time.sleep(0.1)
    pyautogui.click(to_pos)


def main():
    logger = setup_logger()
    logger.info("BejeweledBot started")

    top_left, bottom_right = get_grid_coordinates()
    logger.info(
        "Grid coordinates: top-left=%s, bottom-right=%s", top_left, bottom_right
    )

    last_move = None
    move_count = 0
    start_time = time.time()

    while not keyboard.is_pressed("esc"):
        grid_image = capture_grid(top_left, bottom_right)
        cv2.imshow("Grid Overlay", grid_image)
        cv2.waitKey(100)

        color_grid = build_color_grid(grid_image)

        identified = sum(1 for row in color_grid for c in row if c)
        logger.debug(
            "Grid state (move #%d):\n%s", move_count + 1, format_grid(color_grid)
        )
        logger.debug("Identified %d/%d cells", identified, GRID_SIZE * GRID_SIZE)

        move, score = find_optimal_move(color_grid)

        if move:
            from_row, from_col, to_row, to_col = move

            # Stuck detection: compare with previous move before executing
            is_repeat = last_move == move
            if is_repeat:
                logger.info("Repeat move detected, adding delay")
                time.sleep(0.2)

            move_count += 1
            src_color = color_grid[from_row][from_col]
            dest_color = color_grid[to_row][to_col]
            logger.info(
                "Move #%d: [%d,%d] (%s) -> [%d,%d] (%s) | score=%d%s",
                move_count,
                from_row,
                from_col,
                src_color,
                to_row,
                to_col,
                dest_color,
                score,
                " (repeat)" if is_repeat else "",
            )

            perform_move(top_left, grid_image, from_row, from_col, to_row, to_col)
            last_move = move
        else:
            logger.debug("No valid move found this frame")

    elapsed = time.time() - start_time
    logger.info("Bot stopped. Total moves: %d, Duration: %.1fs", move_count, elapsed)
    cv2.destroyAllWindows()


if __name__ == "__main__":
    main()
