**peregrīnus** m (*genitive* **peregrīnī**); second declension

/pe.reˈɡriː.nus/, [pɛ.rɛˈɡriː.nʊs]

1. Foreigner; traveler
2. (Law) A foreigner who is neither resident nor domiciled in the jurisdiction of the court
3. Tool to parse [NUIX Discover (Ringtail)](https://www.ringtail.com/) exports into legal statements or court books

[![made-with-python](https://img.shields.io/badge/Made%20with-Python-1f425f.svg)](https://www.python.org/)  [![Generic badge](https://img.shields.io/badge/Made%20with-Pandas-yellowgreen)](https://pandas.pydata.org/)  [![GPLv3 license](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://github.com/andre-abadi/peregrinus/blob/master/LICENSE)

# Instructions

1. Place all exports in the `input/` directory
    - Files are processed alphabetically per iteration of this program
2. Run the program: `python peregrinus.py`
3. The program will automatically process the first `.xlsx` file in `input/`
4. At the prompt, choose between converting the file to an Australian legal:
    - Court Book
    - Statement
5. Check `output/` for the finished product (court book or statement)
6. Check `processed/` for the original file that has been moved so the next file in `input/` can be processed
