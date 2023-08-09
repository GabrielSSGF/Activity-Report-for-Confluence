<h1 align="center">
  <br>
  <a href="https://www.atlassian.com/software/confluence"><img src="https://seeklogo.com/images/C/confluence-logo-D9B07137C2-seeklogo.com.png" alt="Jira" width="200"></a>
  <br>
  Activity Report for Confluence
  <br>
</h1>

<h4 align="center">A simple report generator containing the amount of pages created per user </h4>

<p align="center">
  <a href="#key-features">Key Features</a> •
  <a href="#how-to-use">How To Use</a> •
  <a href="#credits">Credits</a>
  
</p>

## Key Features

* Excel File Report
  -  Returns you a *.xlsx* file containing the every user in the atlassian space and the amount of pages created by each of them.
* Page writer
  - Writes the report in a specific page determined by the user.

## How To Use

To clone and run this application, you'll need [Git](https://git-scm.com) and [Python 3](https://www.python.org/downloads/). From your command line:

```bash
# Clone this repository
$ git clone https://github.com/GabrielSSGF/Activity-Report-for-Confluence

# Go into the repository
$ cd Activity-Report-for-Confluence

# Install dependencies
$ pip install pandas openpyxl requests

# Run the app
$ python3 Main.py
```

## Credits

This software uses the following Python libraries:

- [Pandas](https://pandas.pydata.org/) - For data analysis;
- [Openpyxl](https://pandas.pydata.org/) - For *.xlsx* manipulation;
- [Requests](https://pypi.org/project/requests/) - For API requests.
