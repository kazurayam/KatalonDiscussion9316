[Katalon Tips] Multiple Test Cases write data into a single Excel file
======================================================================

This is a [Katalon Studio](https://www.katalon.com/) project for demonstration purpose.
You can clone this out to your PC and run it with your local Katalon Studio.

This project was developed using Katalon Studio v5.5

This project was developed in the hope that it provides a solution to the post into
Katalon Forum:

- [Writing Data in Excel](https://forum.katalon.com/discussion/9316/writing-data-in-excel)

A man posted an example Groovy code which drives Apache POI HSSFWorksheet object to obtain a Excel file. Then another commenter wrote:

>we just write the plain code in the script of our test case or do we need to create a keyword, i have tried many keywords out there but all of them clean my whole spreadsheet and just write the new text katalon send.

I was interested in this question.

The points I found include:

1. he wants to make a single Excel file
2. he wants to update the Excel file multiple times --- possibly 2 or more test cases update the file.

I thought that he seems to have little idea how to implement a Katalon project which supports the above 1 & 2.

So I made this sample project and publish it.

## How to run

TODO

## Specification

TODO

## Design

TODO
