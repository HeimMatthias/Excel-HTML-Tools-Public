# Excel-HTML-Tools-Public
VBA-scripts for rich-text HTML-Export from Excel

VBA-code for *fast* transformation of Excel-cells with rich text formatting (RTF) into html.

This vba-Script transforms the internal XML-represenation of `Range`-content (via `.Value(11)`) via XSLT to html. It represents Cell-Styles as class styles and inline-formatting with tags. The script takes into account that Cell-Styles that are overridden inside cells need to be disabled in the class styles (though with some caveats, see below).
Cell-formats that derive from Cell-presets receive the Preset-name also as a class-attribute, which should make it easy for you to adjust these with your own CSS.

Make sure that the xsl-file is present in the same folder as the Excel-document from which you run the vba-Script.

The XSL-Transformation used can be seen in action [here](https://xsltfiddle.liberty-development.net/jyH9rM8/2), a sample output can be found [here](https://jsfiddle.net/mheim/u5L63cg1/). Note that Excel's internal XML-representation does not specify `xml:space="preserve"`, even though this is required for the transformation to work. The script adds this Attribute before loading.

## Limitations
* Only plain-color cell-background is supported. Patterns could in theory be implemented, but would require svg-background patterns for all possibilites. I suggest to solve this yourself with class-styles. Gradient-backgrounds could easily be implemented in html, but unfortunately they are not represented in Excel's XML-cell representation.
* Cell borders are represented if possible. Some border-styles (`DashDotDot`, `DashDot`, `SlantDashDot`) cannot be represented with CSS and receive their closest match.
* Underline-color of characters does not always take the color of the character. This is because Excel treats this as a character-style, whereas html derives the underline-color always from the `currentColor` of the `U`-node.
* Empty cells (and lines) are skipped over, and are not copied as cells to html. This may result in wrongly aligned columns in the output.
* **Caveat** : Some on/off-Tags, such as `bold` (`underline` or `italics`), are difficult to implement when they derive from cell-styles (i.e. are set for the entire cell with exceptions on some characters added later). Excel checks in the rendering-process if there are individual bold elements inside the cell and then decides to use these tags to render bold passages, effectively disregarding the cell-style that calls for the entire cell to be in bold. This would not work in html.
  
  The XSLT checks whether the first cell using a specific style contains the corresponding tag and omits the css-entry in the style-tag. This is then not applied to all cells using this format.
  
  If this is a problem for you:
  * either only apply the transformation to individual cells and combine them into a unified table yourself, but make sure to untangle conflicting class-styles from the different cell-transformations (which will still receive the same class-name)
  * or adjust the XSLT to no longer remove these CSS-entries from the stylesheet but include checks to insert the negating style-entries into the cell's style-attribute (i.e. `<td style="font-weight:normal"><b>bold</b> and not so bold text combined</td>`. If this is a requirement for you, feel free to open an Issue, and I'll try to help out.
