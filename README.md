# docscoitus
Документачья свадьба
## run params
>go run . <strong>[docx template file path]</strong><strong>[new docx file path]</strong>  <strong>[first template config file path]</strong> <strong>[first xlsx file path]</strong> <strong>[second template config file path]</strong> <strong>[second xlsx file path]</strong>  <strong>third,...</strong> 

test example run:
>go run . testconfig1.txt testtemplate1.docx testresult1.docx test1.xlsx
## template config file <em>(by example examplecfg1.txt)</em>
<em>sheet not necessary, by default takes first (0) sheet</em>
<em>tagnames MUST NOT be the same in various configs</em>
>sheet 1\
>tag1 C1\
>tag2 c10\
>sheet SOMETHING\
>tag3 a1\
>tag4 a2
## tags for replace in docx's text
tag example:
>[[\_yourtagnamehere\_]]

in text:
>texttexttext[[\_tag1\_]]text
