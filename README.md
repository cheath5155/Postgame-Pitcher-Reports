# Postgame Pitcher Report
Overview
Welcome to the Postgame Pitcher Report Tool! This tool is designed to take Trackman CSV data as input and generate a comprehensive PDF report containing various visualizations and a data summary. The generated report includes a breakplot, pitch location chart, release point charts, and a detailed breakdown of the data by pitch type. An example is shown below.

![alt text](https://github.com/cheath5155/Postgame-Pitcher-Reports/blob/master/Example.jpg)

## Installation
1) Clone this repository.
2) Navigate to the project directory.
3) Install dependencies for PostgamePitcherReport.py using pip install.

## Input Requirements
The input CSV file must adhere to the following format:
A Trackman v3 File with Proper Headers

## Usage
1) Ensure PowerPoint is installed on your computer
2) Open PostgamePitcherReport.py and change the pitching team name to the Trackman abbreviation.
3) Run the program and select the CSV or multiple CSVs.
4) The CSVs must be a Trackman V3 formatted CSV or a combined group of Trackman V3 CSVs.
5) The final file will show up in a folder named the date variable. The folder will contain a PDF folder with all PDFs and a FOlder for each player with images for charts and the PowerPoint file.

## Output
The tool will generate a PDF report containing the following visualizations and sections:

1) Breakplot: Visual representation of the break length and break angle of each pitch.
2) Pitch Location Chart: Diagram displaying the location of each pitch on the plate.
3) Release Point Charts: Charts illustrating the release points on the X and Y axes.
4) Data Summary by Pitch Type: Detailed breakdown of the data, including statistics for each pitch type.
