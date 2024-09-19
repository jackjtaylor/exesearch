
# exsearch

exsearch, or Excellent Excel Search, is a module that lets users search across Excel files to 
find relevant cells or data, without having to manually open each one. 

This project was written as a temporary solution, but is now being ported to GitHub and refactored.

## Tech Stack

The project is built fully on Python, using some open source packages like openpyxl to process 
Excel files.


## Installation

To install the project, clone the repository from GitHub.

```bash
  $ git clone https://github.com/jackjtaylor/exsearch
```

Then, open the project and create a local virtual environment with the relevant packages installed.

## Usage

To run a worker or manager, simply invoke the relevant main file.

```python
python3 main.py
```


## Optimisations

The code has been written and refactored to meet PEP8 standards, as well as meet performance standards.

The largest problem with performance, currently, is processing password-locked files. This 
dramatically slows down performance, both by having to enter in the password (I/O slowdowns), and then unlock 
the file.
