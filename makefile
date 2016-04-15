RMD_NAME = example.Rmd

all: rmd docx

alt: rmd docx self

rmd: $(RMD_NAME)
	Rscript --slave -e "rmarkdown::render(input = '$<', output_options = list('self_contained=FALSE'))"

self: $(RMD_NAME)
	Rscript --slave -e "rmarkdown::render(input = '$<', output_options = list('self_contained=TRUE'))"

docx: html2docx.sh
	bash ./html2docx.sh

html2docx.sh:
	wget https://gist.githubusercontent.com/ultinomics/a905343b4ec15e5e212c/raw/eaf107c8a87ca66a1f0ac670edcbc208eecfc4e3/html2docx.sh

