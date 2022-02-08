# xlsxmerge

This program allows you to merge multiple [Microsoft Excel](https://en.wikipedia.org/wiki/Microsoft_Excel) xlsx documents by using a common key identifier.

Let's suppose you have the following tables, stored on different [Microsoft Excel](https://en.wikipedia.org/wiki/Microsoft_Excel) documents.

*File `f1.xlsx`:*

| PersonID | Name     | Surname | Age |
| -------- | -------- | ------- | --- |
| 1        | Mario    | Rossi   | 50  |
| 2        | Carlo    | Bianchi | 40  |
| 3        | Giovanni | Verdi   | 30  |

*File `f2.xlsx`:*

| ID | Address           |
| -- | ----------------- |
| 1  | Rossi street, 1   |
| 2  | Bianchi street, 2 |
| 3  | Verdi street, 3   |

*File `f3.xlsx`:*

| UserId | Nickname | Email              | Country |
| ------ | -------- | ------------------ | ------- |
| 1      | MRossi   | mario@rossi.com    | Italy   |
| 2      | CBianchi | carlo@bianchi.com  | Italy   |
| 3      | GVerdi   | giovanni@verdi.com | Italy   |

At this point, it's possible to unify all information into the same document, by running the following command:

```
python3 xlsxmerge.py -f f1.xlsx f2.xlsx f3.xlsx -k PersonID ID UserId -n ID -o output.xlsx
```

The outcome will be represented by the following Microsoft Excel table.

| ID | Name     | Surname | Age | Address           | Nickname | Email              | Country |
| -- | -------- | ------- | --- | ----------------- | -------- | ------------------ | ------- |
| 1  | Mario    | Rossi   | 50  | Rossi street, 1   | MRossi   | mario@rossi.com    | Italy   |
| 2  | Carlo    | Bianchi | 40  | Bianchi street, 2 | CBianchi | carlo@bianchi.com  | Italy   |
| 3  | Giovanni | Verdi   | 30  | Verdi street, 3   | GVerdi   | giovanni@verdi.com | Italy   |

Thanks to this program, it's possible to unify big data, by keeping 

### Details on usage ###

The launch syntax is the following one:

```
python3 xlsxmerge.py -f <file1.xlsx .. fileN.xlsx> -k <columntitle_file1 .. columntitle_fileN> [-t <tag_file1 .. tag_fileN>] [-o <outputfile.xlsx>]
```

where:
* `-f` specifies the list of files provided in input; each file has to include a single heading line as first line
* `-k` specifies the list of header title used to match the records composing the input files (the equivalent of primary or foreing keys of a [SQL environment](https://en.wikipedia.org/wiki/SQL))
* `-t` specifies, optionally, the list of tags to adopt to distinguish the source of each header (useful to avoid overwriting or to distinguish the source of the property/column/attribute on the final document); in case 
* `-o` specifies, optionally (otherwise, `merged.xlsx` will be considered), the output file name produced

### Installation ###

* Clone the repository:
```
git clone https://github.com/auino/xlsxmerge.git
```
* Enter the program directory:
```
cd xlsxmerge
```
* Install the requirements:
```
pip3 install -r requirements.txt
```

### Contacts ###

You can find me on Twitter as [@auino](https://twitter.com/auino).
