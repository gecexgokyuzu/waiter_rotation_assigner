# Waiter Rotation Scheduler

This Python script helps schedule waiters' rotations in a restaurant by assigning them to different sections of the establishment. The sections have different locations, capacities, and levels of importance.

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)

## Installation

1. Clone this repository to your local machine.
2. Install the required dependencies by running `pip install pandas openpyxl` in your terminal.

## Usage

1. Open the `waiter_rotation_assigner.py` file.
2. Modify the `waiters` and `sections` variables to match your restaurant's needs.
3. Run the script by running `python waiter_rotation_assigner.py` in your terminal.

The script will assign waiters to sections based on the following criteria:

- Waiters can only be assigned to sections that have a capacity greater than 0.
- Waiters cannot be assigned to sections they have already worked in.
- Waiters can only be assigned to sections with the same location as their most recent section if no other sections are available.

The script will output the waiter assignments for the current day and log them in an Excel spreadsheet located in the `./Data` directory.

## Contributing

Contributions are welcome! Please follow the steps below to contribute to this project:

1. Fork this repository to your own GitHub account and clone it to your local machine.
2. Create a new branch and make your changes.
3. Push your changes to your forked repository.
4. Create a pull request to merge your changes into the main repository.

## License

This project is licensed under the MIT License.
