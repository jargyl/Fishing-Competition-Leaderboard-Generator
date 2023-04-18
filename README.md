# Fishing Competition Leaderboard Generator
This is an efficient Python program for generating fishing competition leaderboards, taking into account the number of groups, gender, and weight of fish caught. The program creates both a global and group leaderboard, with automatic segregation of women competitors if needed. The output is an Excel file converted to a Google Sheets link for easy sharing.

## Getting Started
To use this program, you will need to have Python 3 installed on your machine. Clone this repository to your local machine using:

```
git clone https://github.com/<username>/<repository-name>.git
```

### Prerequisites
The following libraries are required to run the program:

* pandas
* gspread
* oauth2client

You can install these libraries using the following command:

```
pip install pandas openpyxl gspread oauth2client
```

### Usage
The program takes a list of participants in the format [("Name", "Gender"), ...]. For example:

```
participants = [("John", "M"), ("Samantha", "F"), ("Mike", "M"), ("Emma", "F")]
```

To divide these participants into groups and create an Excel file, run **group.py**:

```
python group.py
```

The program will create an Excel file named **groups.xlsx**, with the participants divided into groups.

To generate leaderboards based on the amount of fish caught, fill in the amount of KG each participant caught in the tournament in the **groups.xlsx** file. Then, run **leaderboard.py**:

```
python leaderboard.py
```

The program will create a global and group leaderboard, and save them in an Excel file named **leaderboards.xlsx**.

To convert the Excel file to a Google Sheets file, run **sheets.py**. To convert the generated Excel file to a Google Sheets file, you will need to add your Google API credentials to a file named key.json in the project directory. The key.json file should have the following format:

```
{
  "type": "service_account",
  "project_id": "<project-id>",
  "private_key_id": "<private-key-id>",
  "private_key": "<private-key>",
  "client_email": "<client-email>",
  "client_id": "<client-id>",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "<client-cert-url>"
}
```

Replace `<project-id>`, `<private-key-id>`, `<private-key>`, `<client-email>`, `<client-id>`, and `<client-cert-url>` with your own Google API credentials.python sheets.py

Then, run sheets.py to convert the **leaderboards.xlsx** file to a Google Sheets file:
```
python sheets.py
```

Example generated Google Sheet link: https://docs.google.com/spreadsheets/d/1PgkMkrFW9J47Q3U15blDvysVYpsX8-4Bccx-k0T4xl0/


## License
This project is licensed under the MIT License. See the [MIT License](LICENSE) file for more information.
