# Automatic Form 1032 Generation

This repository contains a small Python/QT5 application that automates the tedious task of generating form 1032s. Given an input card inventory Excel file, and a blank template 1032 Excel file, it will produce an output Excel file with one worksheet per group of 5 CACs.   

## Assumptions & Known Issues

This is a somewhat brittle program, and makes certain assumptions, mostly about the layout of the input card inventory file:

- The input inventory file is sorted by batch
- Batch number is in column `J`
- Envelope number is in column `C`
- Proxy number is in column `D`
- Cards per sheet: 5

Beyond that, the other major assumption is about the template file- it basically needs to be run on a very specific template file, otherwise the output behavior is undefined. `¯\_(ツ)_/¯`

Also, Be Ye Warned that there's not a lot in terms of helpful error messages at the moment.

This program also seems to struggle with very large input inventory files, possibly due to limitations in the [openpyxl](https://openpyxl.readthedocs.io/en/stable/) library that I am using to parse Excel workbooks. 

## Installation/Setup

- This has been tested with Python 3.7
- Install necessary dependencies: `pip install -r requirements.txt`
- From there, this project uses the [`fbs`](https://build-system.fman.io) build system to handle packaging (because doing it by hand is a pain and I am lazy). 
    - To run the program, type `fbs run` from the root of the project
    - To cut a release (`.app` on Mac OS X, `.exe` on Windows), run the following commands from the root of the project:
        - `fbs freeze`
        - `fbs installer` 
    - _Note_: This has been tested on Mac OS X and Windows 10; note that on Windows 10 I encountered [some issues](https://github.com/mherrmann/fbs/issues/147#issuecomment-698164639) creating an installer executable.

## Future Features

This was just meant as a quick weekend project, but if anybody ends up ever wanting to use it more, these might be useful features:

- Solving performance issues with large input files
- A more efficient UI for processing multiple batch numbers in one go
- Better error messages
- Handle the case where the input file is not sorted by batch
    - Either by throwing an error message...
    - ... or by just dealing with it internally
- Fewer hard-coded columns, etc.
 