 # html-to-excel-renderer

 ## Install via npm

Npm package contains 2 scripts - postinstall and preuninstall - they are downloading binaries and use go-npm 
(https://www.npmjs.com/package/go-npm) to install them

` npm install -g html-to-excel-renderer`
 
 
 

 ## Install via downloading binaries
 On releases page choose your binary (https://github.com/icewind666/html-to-excel-renderer/releases)
 
 Then download it, unpack and put to accessible location
 (below is an example for linux_amd64)
 
 `wget https://github.com/icewind666/html-to-excel-renderer/releases/download/v1.1.0/html-to-excel-renderer_v1.1.0_linux_amd64.tar.gz`

 `tar -xvf html-to-excel-renderer_v1.1.0_linux_amd64.tar.gz`

 `sudo mv html-to-excel-renderer_v1.1.0_linux_amd64/html-to-excel-renderer /usr/bin`


## Install from source building

 Dependencies: 
 `libxml2-dev`
(`sudo apt-get install libxml2-dev`)
 

`go build -o html-to-excel-renderer .\src\main\main.go`

Run command be like:

`html-to-excel-renderer <template> <data> <output> <batch_size> <debug>`

`html-to-excel-renderer template.hbs data.json result.xslx 5000 0`


**template** - handlebars template file (hbs)

**data** - report data file (json)

**output** - report output

**batch_size** - how match rows to process in one iteration


**debug** - 1 or 0. 

1 - debug mode on. writes rendered.html file with rendered templates.

0 - debug mode off.


 ---
 Constants in code:
  - PIXELS_TO_EXCEL_WIDTH_COEFF (main/utils.go) Transform coeff from pixels to excel width
 
  - PIXELS_TO_EXCEL_HEIGHT_COEFF (main/utils.go) - Transform coeff from pixels to excel height


----
## 3rd party libs

For Handlebars.js template rendering:
**https://github.com/aymerick/raymond**

 For html parsing:
 **https://github.com/jbowtie/gokogiri**
 
 For XLSX generation Excel:
 **github.com/360EntSecGroup-Skylar/excelize/v2**
 
 
 
 
