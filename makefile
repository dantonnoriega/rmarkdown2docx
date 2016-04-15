RMD_NAME = example.Rmd
CSL_FILE = $(CURDIR)/chicago-author-date.csl
BIBLIO_FILE = $(CURDIR)/bibliography.bib

all: rmd docx

alt: rmd docx self

rmd: $(RMD_NAME)
	Rscript --slave -e "rmarkdown::render(input = '$<', output_options = list('self_contained=FALSE'))"

self: $(RMD_NAME)
	Rscript --slave -e "rmarkdown::render(input = '$<', output_options = list('self_contained=TRUE'))"

docx: html2docx.sh
	bash ./html2docx.sh

html2docx.sh:
	wget https://gist.githubusercontent.com/ultinomics/a905343b4ec15e5e212c/raw/b80f18687c913ddebf575438251904ce817cef90/html2docx.sh
