# diffxlsx
diff two text files and output into xlsx spreadsheet

# usage

```
Usage: diffxlsx fromfile tofile output.xlsx

Options: 
  -h, --help            show this help message and exit
  -l LINES, --lines=LINES
                        Set number of context lines (default 0)
  -r REFDIRNAME, --referencedir=REFDIRNAME
                        Directory name for reference content (generate
                        comments from same name file&lines in this directory)
```

# dependencies

+ difflib https://docs.python.org/2/library/difflib.html
+ openpyxl https://bitbucket.org/openpyxl/openpyxl/

# example

For example, there is similar directories orig/ translated/ reviewed/ and 
they have some set of files named identically.

```diffxlsx -r orig translated reviewed out.xlsx```

does compare `translated/*` and `reviewed/*`. 
If there are some different lines in translated and reviewed files, 
they are listed in out.xlsx with comment of `orig/*`


