# Chess Ratings Tracker

![Chess Ratings Tracker Logo](https://upload.wikimedia.org/wikipedia/en/6/67/Lewis_chess_queen_.jpg)

This Python script retrieves chess ratings data from chess.com using the chessdotcom API and updates an Excel file with tables and line charts. The Excel file contains daily ratings data for Blitz, Rapid, and Daily chess games, and a summary sheet showing ratings by month.

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Usage](#usage)
- [File Structure](#file-structure)
- [Example](#example)
- [Author](#author)
- [License](#license)

## Features

- Retrieves daily ratings data from chess.com API.
- Updates Excel file with ratings for Blitz, Rapid, and Daily chess games.
- Generates line charts for each chess category (Blitz, Rapid, Daily) for each day of the current month.
- Updates an overview sheet on the first day of each month with a summary of ratings.

## Requirements

- Python 3.x
- [openpyxl](https://pypi.org/project/openpyxl/)
- [chessdotcom](https://pypi.org/project/chess.com/)

## Usage

1. Install the required Python packages:

    ```bash
    pip install openpyxl chess.com
    ```

2. Update the following information in the Python script:

    - Replace `'sapporoalex'` with your chess.com username.
    - Configure the email address in the `User-Agent` header.
    - Customize the file name and path for the Excel file.

3. Run the script:

    ```bash
    python chess_ratings_tracker.py
    ```

4. Set up the script to run automatically using the task scheduler or cron job everyday at a specific time.

## File Structure

- `Chess Ratings YYYY.xlsx`: Excel file containing ratings data and line charts for each month.
- `chess_ratings_tracker.py`: Python script for retrieving and updating ratings data.

## Example

Here's a sample code snippet from `chess_ratings_tracker.py`:

```Client.request_config['headers']['User-Agent'] = 'My Python Application. Contact me at email@example.com'
data = get_player_stats('sapporoalex', tts=0).json
last_rating_chess_daily = data['stats']['chess_daily']['last']['rating']
last_rating_chess_rapid = data['stats']['chess_rapid']['last']['rating']
last_rating_chess_blitz = data['stats']['chess_blitz']['last']['rating']
```

## Author
Alex McKinley

## License
This project is licensed under the [MIT License](LICENSE).
