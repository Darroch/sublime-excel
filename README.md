# Sublime-excel
Simple syntax file to allow reformatting of Excel formulae in Sublime. Hanldes a number of fairly common Excel Functions to provide colourization and improved readability.

## What is this for:
If you have just met some mega multi-line formula, you will often struggle to edit it nevermind understand it. Once you Copy & Paste it into Sublime you can break it down to examine it there. This gives you the ability to see bracket matching and change the whitespace with tabs, new lines and the like. You can then paste your reformatted formula back into Excel.

## How to Use:
Copy contents to C:\Users\<username>\AppData\Roaming\Sublime Text 3\Packages\User
or for short to %APPDATA%\Sublime Text 3\Packages\User

Then, from Sublime View menu, view->syntax->Excel Formulas. You can also switch in the bottom right of the Sublime Window.

### Included Excel Functions:
- IF,IFS,IFERROR
- AVERAGEIF,AVERAGEIFS,COUNTIF,COUNTIFS,MAXIF,MAXIFS,MINIF,MINIFS,SUMIF,SUMIFS,SUMPRODUCT
- OR,AND,NOT,XOR
- MATCH,INDEX,HLOOKUP,LOOKUP,VLOOKUP,SMALL,LARGE
- CHAR,CHOOSE,CODE,ISTEXT,LEN,SUBSTITUTE,TYPE,UNICHAR
- Numbers
- Operators: &, +,-, /,*,^
- Logical values: TRUE, FALSE

### What is missing:
As you've hopefully realised, we're trying to provide some hope in reformatting complicated Excel formula. This does not handle all Excel Functions, or provide help with Function attributes. However, the ability to add line breaks makes everything more readable, while the syntax highlighting also helps.
No custom colours have been defined, so that they do not clash with your preferred 'Color Scheme'.
- [ ] Improved handling of Table names
- [x] Highlighting volatile Functions: CELL,INDIRECT,INFO,NOW,OFFSET,RAND,TODAY

## Additional Resources:
In creating this Syntax file, the following resources were used:

1. https://www.sublimetext.com/docs/3/syntax.html
2. http://docs.sublimetext.info/en/latest/extensibility/syntaxdefs.html
