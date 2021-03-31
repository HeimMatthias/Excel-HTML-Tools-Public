# Excel-HTML-Tools
VBA-scripts for rich-text HTML-Export from Excel

VBA-code for *fast* transformation of Excel-cells with rich text formatting (RTF) into html.

This vba-Script transforms the internal XML-represenation of `Range`-content (via `.Value(11)`) via XSLT to html. It represents Cell-Styles as class styles and inline-formatting with tags. The script takes into account that Cell-Styles that are overridden inside cells need to be disabled in the class styles (though with some caveats, see below).
Cell-formats that derive from Cell-presets receive the Preset-name also as a class-attribute, which should make it easy for you to adjust these with your own CSS.

Make sure that the xsl-file is present in the same folder as the Excel-document from which you run the vba-Script.

The XSL-Transformation used can be seen in action [here](https://xsltfiddle.liberty-development.net/jyH9rM8/2), a sample output can be found [here](https://jsfiddle.net/mheim/u5L63cg1/). Note that Excel's internal XML-representation does not specify `xml:space="preserve"`, even though this is required for the transformation to work. The script adds this Attribute before loading.

## How do I use this?
Simply load the script [Excel-RTF-to-HTML-via-XSLT.vbs](https://github.com/HeimMatthias/Excel-HTML-Tools-Public/blob/main/Excel-RTF-to-HTML-via-XSLT.vbs) into your Excel sheet. Or download the [Macro-enabled Excel-file](https://github.com/HeimMatthias/Excel-HTML-Tools-Public/blob/main/Excel-Macro-turn-RTF-formatted-cells-into-html.xlsm) from this repository. Save the XSL-Stylesheet you intend to use in the same folder (more on this below).

*If you are a programmer* access the `function fnConvertXML2HTML()` in your own code, by passing on the xml from an Excel range. E.g. `Dim html As String: html = fnConvertXML2HTML(Range("A1:H8").Value(11))`.

*If you are a user* simply use one of the helper-functions. Select the content you want to export as html and then [run one of the following macros](https://support.microsoft.com/en-us/office/run-a-macro-5e855fd2-02d1-45f5-90a3-50e645fe3155).
* `copySelectionAsHTML`: this will copy your selection as html to the clipboard (note that you will need to insert this code into an html-boilerplate, since `html`- and `body`- tag are missing).
* `saveSelectionAsHTML`: this will prompt you with a "Save As"-dialog where you can enter a filename for your html-file.
* `turnSelectionRichTextIntoHtml`: this will turn each of the cells in the selection individually into html. *Beware that this will overwrite your data.* Use this if you plan to export the html of individual cells into a database, csv, conversion tool, etc.
* `turnSelectionNumbersIntoText`: I **strongly advise** you to run this helper script before you export your html. *Beware that this will overwrite your data.* Doing so will allow you to export formatted numbers and empty cells. Only do not run this beforehand, if your Excel contains only formatted text and missing cells are not a problem for you.

## What transformations are available?
There are three different transformations available:
* `excel-complex-style-transformation.xsl` [default]: This transformation will retain all of Excel's internal style sheets and use them to format the output. Only formatting inside the cell will be done with additional html-tags. If a cell deviates from its style-definition, the exceptions are defined in the cell's style attribute.
* `excel-full-transformation.xsl` : This will only retain styles as a style sheet that concern the outer appearance of cells (i.e. borders, background-color, alignment). But all text-/character-/font-level formatting will be translated into html-tags
* `excel-simple-style-transformation.xsl` : This transformation retains the style-sheets, but attempts to adjust the styles based on the overriding formats from the first cell that uses the style. This will produce cleaner output than the default transformation because the style-attribute of the cells are not used to define exceptions. However, it cannot resolve conflicts from multiple cells that use the same style. So beware when using it to export more than individual cells.
* 
If you use a different XSLT than the default `excel-complex-style-transformation.xsl`, you need to adjust the constant `xsltConversionFile` in the second line of the script.

## Limitations
There are some aspects of Excel's formatting options that cannot be reproduced with html and css. There are also some border-line cases that have not been implemented. Please check the [limitations in the issues](https://github.com/HeimMatthias/Excel-HTML-Tools-Public/issues?q=is%3Aissue+is%3Aopen+Limitations)

There are, however, two **important limitations** to the html conversion:
* **Empty cells** (and lines) **are skipped over**, and are not copied as cells to html. This may result in wrongly aligned columns in the output.
* **Formatted numbers and formula** see their values copied without formatting.

**Workaround:** Use the included script `turnSelectionNumbersIntoText()` to prepare your worksheet before exporting to html. This will turn formatted numbers and formula into text, but this is a destructive solution that cannot be undone. Be careful when saving your document.

## but what about Excel's html-Export
By all means use Excel's html-Export if you need the full file exported as html and do not need to programmatically access the html from VBA. Excel's export separates css and html; but more importantly, it does not produce html that could easily be re-used or parsed. In Excel's native export all font and character formats are applied via `font`-tags with individual classnames.

This XSL-transformation here will retain cleaner html tags such as `<i>, <b>, <del>, <sup>, <sub>, <u>` and processes the remaining formats inside `<span>`-tags with clean style-attributes. This means, that you can easily transfer the formatted output from Excel to another tool if you rely on the XSL-transformation of this tool, but not when you rely on Excel's output.

If you need to convert a full worksheet, including several tables and need all formats to work, but do not intend to re-use or parse the content outside the exported data, Excel's native export should be the way forward for you.
