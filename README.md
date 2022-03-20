# txt-to-excel-converter

## How to use

0. prepare an excel file you want to update
1. create a md file `example.md`

```md
file: example.xlsx

## Sheet1

cell equals: This is A1
add-r: add to B1
and to B1 again
add-r: add to C1
cell starts: This is C2,
add-l: add to B2
add-b: add to B3
```

2. run cmd `Excel.Editor.exe example.md`

## Commands

### for settings

|Command|Explaination|
|---|---|
|`file`|file path of an excel file|
|`output`|file path of the new excel file|
|`params`|other parameters|
|`fill`|specify blank columns, when `user-title` is set|

### for `params`

all parameters in this section can be appended to `params:` line.

|Parameter|Explaination|
|---|---|
|`use-title`|use title to select column index|

### for cells

|Command|Explaination|
|---|---|
|`#`, `##`|specify a sheet name|
|`cell equals`|find row, column indices by comparing `cell == text`|
|`cell starts`|find row, column indices by comparing `cell starts with text`|
|`add-r`|move to the right cell, and add text|
|`add-l`|move to the left cell, and add text|
|`add-b`|move to the bottom cell, and add text|
|`add-t`|move to the up cell, and add text|
