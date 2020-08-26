**peregrīnus** m (*genitive* **peregrīnī**); second declension

1. foreigner; traveler
2. (law) a foreigner who is neither resident nor domiciled in the jurisdiction of the court

[![made-with-python](https://img.shields.io/badge/Made%20with-Python-1f425f.svg)](https://www.python.org/) [![GPLv3 license](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://github.com/andre-abadi/peregrinus/blob/master/LICENSE)

# Introduction 

Tool to parse [NUIX Discover (Ringtail)](https://www.ringtail.com/) exports into:
    - Paragraph text suitable for use in legal statements
    - A court book for use in legal proceedings

# Instructions

1. Place all exports in the `input/` directory
    - Files are processed alphabetically per iteration of this program
2. Run the program: `python peregrinus.py`
3. The program will automatically process the first `.xlsx` file in `input/`
4. At the prompt, choose between converting the file to a:
    - Court Book
    - Statement
5. Check `output/` for the finished product (court book or statement)
6. Check `processed/` for the original file that has been moved so the next file in `input/` can be processed
