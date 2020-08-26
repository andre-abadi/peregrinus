# Peregrinus

**peregrīnus** m (**genitive*** *peregrīnī**); second declension
1. foreigner; traveler
2. (law) a foreigner who is neither resident nor domiciled in the jurisdiction of the court

# Introduction 

Tool to parse [NUIX Discover (Ringtail)](https://www.ringtail.com/) exports into:
  - Paragraph text suitable for use in legal statements
  - A court book for use in legal proceedings

# Instructions

1. Place all exports in the `input/` directory
  - Files are processed alphabetically per iteration of this program
2. Run the program: `python peregrine.py`
3. The program will automatically select first `.xlsx` file in `input/`
4. Choose between converting the file to a:
  - Court Book
  - Statement
5. Check `output/` for the finished product
6. Check `processed/` for the original file
