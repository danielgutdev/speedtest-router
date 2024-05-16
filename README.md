# Speedtest and Router Status Checker

This script performs internet speed tests using Speedtest.net and checks the status of your router at regular intervals. The results are saved to an Excel file.

## Requirements

- Python 3.x
- Google Chrome browser

## Installation

### Windows

1. **Clone the repository:**
    ```sh
    git clone https://github.com/yourusername/speedtest-router.git
    cd speedtest-router
    ```

2. **Create and activate a virtual environment (optional but recommended):**
    ```sh
    python -m venv venv
    venv\Scripts\activate
    ```

3. **Install the required packages:**
    ```sh
    pip install -r requirements.txt
    ```

### Linux

1. **Clone the repository:**
    ```sh
    git clone https://github.com/yourusername/speedtest-router.git
    cd speedtest-router
    ```

2. **Create and activate a virtual environment (optional but recommended):**
    ```sh
    python3 -m venv venv
    source venv/bin/activate
    ```

3. **Install the required packages:**
    ```sh
    pip install -r requirements.txt
    ```

## Configuration

1. **Update the `config.ini` file with your settings:**
    ```ini
    [Settings]
    Interval = 60
    RouterWebsite = http://192.168.1.1/
    LoopCount = 10
    ```

## Usage

1. **Run the script with default configuration:**
    ```sh
    python speedtest.py
    ```

2. **Override the interval and loop count via command-line arguments:**
    ```sh
    python speedtest.py --interval 120 --loops 5
    ```

## Example

To run the script with a 2-minute interval and 5 iterations:
```sh
python speedtest.py --interval 120 --loops 5