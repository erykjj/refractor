# refractor

Commandline/terminal **reference extractor** utility that finds scriptures and publication references[^*]. Works with both text files and *.docx* (MS Word) documents. Multiple languages are supported. Publication references only parsed within `()` and `[]`. You will get a list that you can paste into the search box on [*Watchtower ONLINE LIBRARY*](https://wol.jw.org)[^#], or a link to see your results directly.

## Downloads

- [Linux (x86_64)](https://github.com/erykjj/refractor/releases/latest/download/refractor_linux_x86_64.tgz)
- [Linux (ARM64)](https://github.com/erykjj/refractor/releases/latest/download/refractor_linux_arm64.tgz)

- [macOS (universal)](https://github.com/erykjj/refractor/releases/latest/download/refractor_macos.tar.gz)

- [Windows (amd64)](https://github.com/erykjj/refractor/releases/latest/download/refractor_windows_amd64.zip)
- [Windows (ARM64)](https://github.com/erykjj/refractor/releases/latest/download/refractor_windows_arm64.zip)

## Usage

```
 Usage: refractor [-h | -v | -l] | [-r] [-s] -c:code <infile>

 Options:
   -h, --help                      Show this help message and exit
   -v, --version                   Show the version and exit

   -c:<code>, --code=<code>        Language code (en by default)
   -l, --list                      List supported languages

   -r, --references                Output publication references
   -s, --scriptures                Output scriptures (if neither -r or
                                     -s is provided, both are enabled)

 <infile>                          File to process (docx or text)
```

## Examples

Linux/macOS: `./refractor -c:es bosquejo.docx`

Windows: `refractor.exe -c:en talk_outline.txt`

____
[![Static Badge](https://img.shields.io/badge/releases-orange?style=plastic&logo=rss&logoColor=orange&color=black)](https://github.com/erykjj/refractor/releases.atom)

By using this software you agree to abide by the terms of its [License](https://github.com/erykjj/refractor#License-1-ov-file).

Feel free to get in touch and post any [issues and/or suggestions](https://github.com/erykjj/refractor/issues).

____
#### Footnotes:
[^*]: For publications of Jehovah's Witnesses
[^#]: Copyright *Watch Tower Bible and Tract Society of Pennsylvania*
