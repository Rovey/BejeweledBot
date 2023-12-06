import cv2
import numpy as np
import pyautogui
import time
import keyboard

# Define colors
COLORS = {
    (206, 165, 33): "blue",
    (41, 115, 16): "green",
    (16, 49, 239): "red",
    (206, 70, 196): "purple",
    (24, 123, 255): "orange",
    (24, 222, 255): "yellow",
    (211, 211, 211): "white",
}

# Initialize an 8x8 grid map to keep track of found colors
color_grid = [["" for _ in range(8)] for _ in range(8)]


def get_grid_coordinates():
    """
    Get the top-left and bottom-right coordinates of the game grid by prompting the user to
    click on the corners.
    """
    input("Move your mouse to the top-left corner of the game grid and press Enter.")
    top_left_corner = pyautogui.position()
    input("Move your mouse to the bottom-right corner of the game grid and press Enter.")
    bottom_right_corner = pyautogui.position()
    return top_left_corner, bottom_right_corner


def capture_grid(top_left_corner, bottom_right_corner):
    """
    Capture the game grid within the specified region and return the grid image.
    """
    screenshot = pyautogui.screenshot(
        region=(
            top_left_corner[0],
            top_left_corner[1],
            bottom_right_corner[0] - top_left_corner[0],
            bottom_right_corner[1] - top_left_corner[1],
        )
    )
    grid_img = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

    # Draw grid lines
    cell_width = grid_img.shape[1] // 8
    cell_height = grid_img.shape[0] // 8
    for i in range(1, 8):
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


def mark_cell(grid_img, roww, coll):
    """
    Mark the color in the grid map based on the specified row and column.
    """
    global color_grid

    cell_width = grid_img.shape[1] // 8
    cell_height = grid_img.shape[0] // 8
    cell = grid_img[
        roww * cell_height : (roww + 1) * cell_height,
        coll * cell_width : (coll + 1) * cell_width,
    ]

    for color, name in COLORS.items():
        # Create a binary mask for the specified color within a tolerance
        mask = cv2.inRange(cell, np.array(color) - 3, np.array(color) + 3)

        # Check if any non-zero pixels in the mask
        if cv2.countNonZero(mask) > 0:
            color_grid[roww][coll] = name
            return


def find_optimal_move():
    """Tries to find optimal position to move a gem to for the highest score."""
    best_move = None
    best_score = -float("inf")

    for outer_row in range(7, -1, -1):
        for inner_col in range(8):
            if not color_grid[outer_row][inner_col]:
                continue

            # Check for a potential move to the right
            if inner_col < 7:
                color_grid[outer_row][inner_col], color_grid[outer_row][inner_col + 1] = (
                    color_grid[outer_row][inner_col + 1],
                    color_grid[outer_row][inner_col],
                )
                score = evaluate_state(color_grid)
                color_grid[outer_row][inner_col], color_grid[outer_row][inner_col + 1] = (
                    color_grid[outer_row][inner_col + 1],
                    color_grid[outer_row][inner_col],
                )

                if score > best_score:
                    best_score = score
                    best_move = (outer_row, inner_col, outer_row, inner_col + 1)

            # Check for a potential move down
            if outer_row < 7:
                color_grid[outer_row][inner_col], color_grid[outer_row + 1][inner_col] = (
                    color_grid[outer_row + 1][inner_col],
                    color_grid[outer_row][inner_col],
                )
                score = evaluate_state(color_grid)
                color_grid[outer_row][inner_col], color_grid[outer_row + 1][inner_col] = (
                    color_grid[outer_row + 1][inner_col],
                    color_grid[outer_row][inner_col],
                )

                if score > best_score:
                    best_score = score
                    best_move = (outer_row, inner_col, outer_row + 1, inner_col)

    return best_move


def evaluate_state(grid):
    """Checks the state of the grid and returns the score."""
    score = 0

    # Evaluate based on the number of valid consecutive colors in a row
    for row in range(8):
        for col in range(8):
            color = grid[row][col]
            if color:
                # Check horizontally to the right
                consecutive_right = 1
                for i in range(1, 5):
                    if col + i < 8 and grid[row][col + i] == color:
                        consecutive_right += 1
                    else:
                        break

                # Check vertically down
                consecutive_down = 1
                for i in range(1, 5):
                    if row + i < 8 and grid[row + i][col] == color:
                        consecutive_down += 1
                    else:
                        break

                # Update the score based on valid consecutive colors
                if consecutive_right >= 3:
                    score += consecutive_right
                if consecutive_down >= 3:
                    score += consecutive_down

    return score


# Function to perform a move
def perform_move(src_row, src_col, dest_row, dest_col):
    """
    Perform a move from the specified source to the destination coordinates.
    """
    # Calculate the positions in screen coordinates
    from_position = (
        top_left[0]
        + src_col * (grid_image.shape[1] // 8)
        + (grid_image.shape[1] // 16),
        top_left[1]
        + src_row * (grid_image.shape[0] // 8)
        + (grid_image.shape[0] // 16),
    )

    to_position = (
        top_left[0]
        + dest_col * (grid_image.shape[1] // 8)
        + (grid_image.shape[1] // 16),
        top_left[1]
        + dest_row * (grid_image.shape[0] // 8)
        + (grid_image.shape[0] // 16),
    )

    # Simulate mouse click to perform the move
    pyautogui.click(from_position)
    time.sleep(0.1)  # Add a small delay between clicks
    pyautogui.click(to_position)


# Get grid coordinates once before entering the main loop
top_left, bottom_right = get_grid_coordinates()

# Main loop
while True:
    grid_image = capture_grid(top_left, bottom_right)

    cv2.imshow("Grid Overlay", grid_image)
    cv2.waitKey(100)

    # Loop through each cell and mark the color
    for row in range(8):
        for col in range(8):
            mark_cell(grid_image, row, col)

    # Search for an optimal move
    move_coordinates = find_optimal_move()

    if move_coordinates:
        # Perform the move using the perform_move function
        from_row, from_col, to_row, to_col = move_coordinates
        print(f"Performing move: [{from_row}, {from_col}] -> [{to_row}, {to_col}]")
        perform_move(from_row, from_col, to_row, to_col)

    # Check for the "Escape" key press to exit the script
    if keyboard.is_pressed("esc"):
        break
