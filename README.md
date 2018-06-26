# xlsx2markdown
Excel ファイル(.xlsx)の指定シート内容を Markdown に落とす Python スクリ。

## Requirement
- Windows 7+
- Python 3.6+
  - xlrd==1.1.0

## Demo

### (Before) Excel file
sample.xlsx が以下の三シートを含んでいるとします(xlsxファイル自体は個人情報が含まれてしまうため同梱していません)。

!(sheet1)[xlsx2markdown_demo1.jpg]

!(sheet2)[xlsx2markdown_demo2.jpg]

!(sheet3)[xlsx2markdown_demo3.jpg]

### (After) Conversion
以下のようにして変換を行います。

```
$ python converter.py -t sample.xlsx 0 1 2
[1/3]...
[2/3]...
[3/3]...
fin.
```

すると以下のような Markdown ファイルが出来上がります。

- [sheet1](demo1.md)
- [sheet2](demo2.md)
- [sheet3](demo3.md)

中身について、たとえば sheet1 の分を以下に取り上げてみます。

```
# Line 1

## 1 - A
A1

## 1 - B
B1

## 1 - C
C1

# Line 2

## 2 - A
A2

## 2 - B
B2

...
```

一般化して書くと、以下のようなフォーマットになっています。

```
# Line 1

## 1 - A
一行目A列の内容

## 1 - B
一行目B列の内容
...

# Line 2

## 2 - A
二行目A列の内容

## 2 - B
二行目B列の内容
...

```

## Installation
- `pip instal xlrd`
- `git clone https://github.com/stakiran/xlsx2markdown`
- `cd xlsx2markdown`
- `python converter.py (OPTIONS...)`

## License
[MIT License](LICENSE)

## Author
[stakiran](https://github.com/stakiran)
