# docscoitus
Документачья свадьба
## run params
>go run . <strong>[template config file path]</strong> <strong>[docx template file path]</strong> <strong>[new docx file path]</strong> <strong>[xlsx file path]</strong>

test example run:
>go run . testconfig1.txt testtemplate1.docx testresult1.docx test1.xlsx
## template config file <em>(by example testconfig1.txt)</em>
<em>sheet not necessary, by default takes first (0) sheet</em>
>sheet 1
>tag1 C1
>tag2 c10
## tags for replace in docx's text
tag example:
>[[\_yourtagnamehere\_]]

in text:
>texttexttext[[\_tag1\_]]text
