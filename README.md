# introdoc

Introdoc (from `introspection` and `document`) — a tool for generating `.docx` documents, containing introspected tables from Microsoft SQL Server databases.

Created for needs of students in Russian colleges, who have to attach tables with structure of databases to their reports for practical works and course projects. Basically, when you have 10+ tables in your database, it's a pain to write them all manually. So, there it is.

## Installation

1. Clone this repository with `git clone https://github.com/GD-alt/introdoc`.
2. Install dependencies with `pip install -r requirements.txt`. If you're using poetry, you can run `poetry install` instead.
3. Now you're able to run the script with `py -m introdoc`. If you're a happy Linux user, you can use `python3 -m introdoc` instead.

## Troubleshooting

This script uses `pyodbc` to connect to SQL Server. You might need to install ODBC drivers for your system. If you're using Windows, you can download them from [Microsoft](https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver16#download-for-windows) site. If you're using Linux… well, you're on your own, I don't use it in everyday life. Good luck!

## Usage

Introdoc is a command-line tool, which provides hella load of options. Here's a quick overview of them:

|   Long name   | Short name |                         Description                         |            Default            |
|:-------------:|:----------:|:-----------------------------------------------------------:|:-----------------------------:|
|  `--output`   |    `-o`    |                    Output document name                     |          output.docx          |
| `--database`  |    `-d`    |                   Database to introspect                    |            master             |
|  `--server`   |    `-s`    |                         Server name                         |    (localdb)\\mssqllocaldb    |
|  `--driver`   |    `-D`    |                         Driver name                         | ODBC Driver 17 for SQL Server |
|  `--tables`   |    `-t`    |           Tables to introspect, divided by comma            |             None              | 
| `--language`  |    `-l`    |        Language to use in tables (`ru`, `en`, `de`)         |              en               |
|  `--headers`  |    `-h`    |        If to include table headers into table (flag)        |             False             |
| `--inullable` |    `-N`    |     If to include `Nullable?` column into table (flag)      |             False             |
| `--onatural`  |    `-n`    | If included, `Nullable column` will contain `Yes` and `No`s |             False             |
| `--sections`  |    `-s`    |        If to separate tables with their names (flag)        |             False             |

## Examples

- `py -m introdoc` will introspect all tables from `master` database on `(localdb)\\mssqllocaldb` server, and save them into `output.docx` file. Tables will be in English, without headers, `Nullable?` column, `Yes` and `No` values in `Nullable?` column, and not separated with their names.
- `py -m introdoc -d AdventureWorks -s localhost -D 'ODBC Driver 17 for SQL Server' -t HumanResources.Department,HumanResources.Employee -l ru -h -N -n -s` will introspect `HumanResources.Department` and `HumanResources.Employee` tables from `AdventureWorks` database on `localhost` server, using `ODBC Driver 17 for SQL Server` driver, and save them into `output.docx` file. Tables will be in Russian, with headers, `Nullable?` column, `Yes` and `No` values in `Nullable?` column, and separated with their names.
- `py -m introdoc -d 10210797 -S` (my usecase) will introspect all tables from `10210797` database on `(localdb)\\mssqllocaldb` server, and save them into `output.docx` file. Tables will be in English, without headers, `Nullable?` column, `Yes` and `No` values in `Nullable?` column, and not separated with their names.