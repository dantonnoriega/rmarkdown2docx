MAIN_DIR = ~/Github/rmarkdown2docx/
FILE_BASE_NAME = example
CSL_FILE = $(MAIN_DIR)/chicago-author-date.csl
BIBLIO_FILE = $(MAIN_DIR)/bibliography.bib

all: rmd docx

alt: rmd md docx

rmd:
	Rscript --slave -e "rmarkdown::render(input = '$(FILE_BASE_NAME).Rmd')"

md:
	pandoc -f markdown+simple_tables+table_captions+yaml_metadata_block -t html $(MAIN_DIR)/$(FILE_BASE_NAME).md -s -S -o $(FILE_BASE_NAME).html --filter pandoc-citeproc --csl $(CSL_FILE) --bibliography $(BIBLIO_FILE)

docx:
	sh ./html2docx.sh