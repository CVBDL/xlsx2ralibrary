# xlsx2ralibrary

RaLibrary command line import tool.

> The input file must be closed when running this tool.
> Because we will write logs to this file.

## Prerequisites

* [Python3.x](https://www.python.org/ftp/python/3.6.3/python-3.6.3.exe)

## Excel content format

The first row should be a header row.
And this row will be ignored when creating books.

| ISBN          | Code  | Book Name            |
| ------------- | ----- | -------------------- |
| 9780596008031 | P501  | Designing Interfaces |

## Examples

```sh
# show help
xlsx2ralibrary.py --help

# import data
python3 xlsx2ralibrary --user-name=username --password=pwd --path=./xlsx2ralibrary/books.xlsx
```
