# Docx-Equations
Converts equations in a Word (.docx) document to LaTeX format and exports the content as an HTML file.
</br></br>

## Background


A Word (.docx) file is essentially a ZIP file. Upon unzipping, multiple documents are extracted, with the main content typically located in `word/document.xml`.
```    
        $: unzip equation.docx
        $: tree
        ├── equation.docx
        ├── [Content_Types].xml
        ├── docProps
        │   ├── app.xml
        │   └── core.xml
        ├── equation.docx
        ├── _rels
        └── word
            ├── document.xml
            ├── fontTable.xml
            ├── _rels
            │   └── document.xml.rels
            └── styles.xml
 ```   

### There are three ways to insert a equation into word file:

- eq field 
- omml
- MathType 

All of them are stored as markup inside docx file. You can easily find them at **word/document.xml**.

```
    <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <m:sSup>
            <m:e>
                <m:d>
                    <m:dPr>
                        <m:begChr m:val="("/>
                        <m:endChr m:val=")"/>
                    </m:dPr>
                    <m:e>
                        <m:r>
                            <w:rPr>
                                <w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>
                            </w:rPr>
                            <m:t xml:space="preserve">x</m:t>
                        </m:r>
                    .......
                    </m:e>
                </m:sSup>
            </m:e>
        </m:nary>
    </m:oMath>
```
So we can just extract the equation markup and convert it into LaTeX.
