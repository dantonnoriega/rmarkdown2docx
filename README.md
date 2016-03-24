# Description

This set of scripts help convert the output of `Rmd` or `md` files to `docx` files. It is done by creating a clean `html` file, opening and saving the `html` to `docx` in Microsoft Word.

The workhorse scripts are:

- `makefile`
- `html2docx.sh`

These can be changed to convert any `Rmd` file but with some caveats outlined below. In this repo, the scripts convert `example.Rmd` to `example.docx`.

The output files are:

- `example.html`
- `example.md`
- `example.docx`

Optional support files are:

- `chicago-author-date.csl`
- `bibliography.bib`

These files are listed to show that folks can cite references (useful for academics).

# Requirements

1. Microsoft Word for Mac in your Applications folder. This has been built and testing using Microsfot Word for Mac, Version `15.20`.
2. `R` with packages `rmarkdown` and `knitr`.
3. An understanding of how to use GNU Make and terminal commands.

# How to Use

There are two options for going from `Rmd` to `docx`.

If you look at the `makefile`, the first is option is `all` and the second is `alt`.

## Option 1: `all`

Although one can knit an `html` file from a `Rmd` file---letting knitr run the `pandoc` step---there is a caveat.

> yaml option `self_contained` **must** be `false`. Otherwise, Microsoft Word will crash during the `html` to `docx` conversion. Keep in mind that this is the default option for `rmarkdown`.

Also, it doesn't matter if `keep_md: true`. I just prefer to have the `.md` regardless of what I'm doing. Basically, you want to make sure the final `html` file produce is as simple and clean as possible.


## Option 2: `alt`

You cannot convert a self contained (aka standalone) `html` file to a `docx`. (At least I've found that it always crashes.) If you want the option to have a standalone `html` file, then you must enable the following yaml options in the `Rmd` file:

1. `self_contained: true`
2. `keep_md: true`

One then runs, in order, the following make commands

1. `make alt`
2. `make rmd`

`make alt` will knit a standalone `html` file but also produce a `.md` file. It will then take the `.md` file and produce a clean (not self contained) `html` file using `pandoc`, replacing the standalone `html` file. The `.sh` script will then take the clean `html` file and create a `docx` file.

`make rmd` will then knit the standalone `html` file, replacing the clean `html` file.

It's a little hackish, but it allows you to create a `docx` AND keep a standalone `html` file.

**Note**: if you use option 2 and set the yaml date option to `'`r format(Sys.time(), "%B %d, %Y")`'`, it will NOT convert. Only option 1 does this. Haven't quite figured out how to do that for option 2. So, if you use option 2, you'll have to fill in a date manually (e.g. `date: March 1, 2016`).


# Scripts `makefile` and `html2docx.sh`

A few things to note in these scripts. In it's current state, both are set to run out directory `~/Github/rmarkdown2docx/`. That is, a path searching *my* home directory and pathing to directory `rmarkdown2docx` (which is contained in directory `Github`). Apologies for being pedantic, but pathing tends to trip up novice users.

If you download this repo, change the main directory path in both the `makefile` and `html2docx.sh` files and then the makefile should run.

More specifically...

## `makefile`

In `makefile` change variable `MAIN_DIR` to the path that contains the files in this repo (whether you fork or download):

```
MAIN_DIR = ~/Github/rmarkdown2docx/ # CHANGE THIS. "~" SHOULD BE FINE.
FILE_BASE_NAME = example
CSL_FILE = $(MAIN_DIR)/chicago-author-date.csl # THIS CAN BE ANY PATH TO ANY CSL FILE
BIBLIO_FILE = $(MAIN_DIR)/bibliography.bib # SAME FOR ANY BIB FILE
```

## `html2docx.sh`

In `html2docx.sh` change variable `base_folder` to the path that contains the `html` file you want to convert. If you change path to the fork or download of this repo, then it will search for file `example.html` and convert it to `example.docx`.

```
set base_folder to "/Users/dnoriega/Github/rmarkdown2docx/" # CANNOT use "~". MUST BE FULL PATH. MUST END IN "/"
set file_in to base_folder & "example.html"
set file_out to base_folder & "example.docx"
```

When you first run the script, do not worry if Word asks for permissions. Once you give Word access to the folder and files, it should run just fine.

# Thanks

A special thanks to [Andrew Heiss](http://github.com/andrewheiss), from whom I've learned almost all I know about makefiles and converting markdown files to docx files.



