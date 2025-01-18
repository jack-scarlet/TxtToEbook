# TxtToEbook

A quick tool that I've made to generate .epub ebook file from txt.

The tool works by inclusing each .txt as an chapter.
If you have a single txt with multiple chapters, change the extesion to .btxt

The .btxt files will be splited in multiple .txt files including a single chapter each.

The chapter content can have a [image] (up two per chapter), this tag will search for a image with the chapter number in the folder img for example, if [image] is found in chapter 20, the tool will find for the file img\20.jpg, if a second [image] tag is in the same chapter, it will look for the file img\20-2.jpg

Sample Input Directory Structure
```
img\
img\20.jpg
chapters.btxt
cover.jpg
```
Alternative structure
```
img\1.jpg
img\2.jpg
1.txt
2.txt
3.txt
4.txt
...
cover.jpg
```
Now you can add images directly to the .btxt, just start the line with the <image> tag
```
text
text
<image>1.jpg
text
text
<image>2.jpg
text
text
...
```

Support for declaring title tags with h1, h2... has also been added.
```
text
text
<h1>Hello World
text
text
<h2>Welcome text
text
text
...
```
*When using the h1 and h2 tags, the ebook index (toc.ncx) will use it as a reference.

After the files are ready, just drag&drop the root directory to the TxtToEbook.exe and the epub will be created.


### .btxt files
The .btxt files are just plain text files with multiple chapters that is delimited by the chapter number
```
Chapter 1

This is the story of a man named Stanley

Chapter 2

Stanley worked for a company in a big building where he was Employee Number 427.
```

