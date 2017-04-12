# BEFORE docx conversion

For all footnote refs, make footnoteref the only class for that span (currently they come in with class="superscriptchars footnoteref").

For all footnotes, convert parent spans to paras.

Correctly formatted HTML tables will be converted correctly by the PS1.

# INCOMING footnote format

```
<w:p><w:pPr><w:divId w:val="693382070"/><w:rPr><w:rFonts w:eastAsia="Times New Roman"/></w:rPr></w:pPr><w:r><w:rPr><w:rStyle w:val="footnotetext"/><w:rFonts w:eastAsia="Times New Roman"/></w:rPr><w:t xml:space="preserve">Well, that is, at least, how one member of an online discussion group humorously put it in a forum regarding the matter (Scrap-Lord, </w:t></w:r><w:r><w:rPr><w:rStyle w:val="spanhyperlinkurl"/><w:rFonts w:eastAsia="Times New Roman"/></w:rPr><w:t>http://forum.deviantart.com/devart/general/2226526</w:t></w:r><w:r><w:rPr><w:rStyle w:val="footnotetext"/><w:rFonts w:eastAsia="Times New Roman"/></w:rPr><w:t>).</w:t></w:r></w:p>
```

# ADDING FOOTNOTES

First two children are the separators:

```
<w:footnote w:type="separator" w:id="-1"><w:p w14:paraId="68D87997" w14:textId="77777777" w:rsidR="00A51F92" w:rsidRDefault="00A51F92"><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>
<w:footnote w:type="continuationSeparator" w:id="0"><w:p w14:paraId="281652F2" w14:textId="77777777" w:rsidR="00A51F92" w:rsidRDefault="00A51F92"><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
```

Then the real footnotes:

```
<w:footnote w:id="1"><w:p w14:paraId="012E3F1F" w14:textId="77777777" w:rsidR="00A51F92" w:rsidRDefault="00A51F92"><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Baritone </w:t></w:r><w:r w:rsidRPr="00DE2E06"><w:rPr><w:rStyle w:val="spanitaliccharactersital"/></w:rPr><w:t xml:space="preserve">con squillo. </w:t></w:r><w:r><w:t>A voice to be heard, in every sense: clarion-like over the din of skirmish, agonies, and war-cries; and as daunting to his own as to the enemy.</w:t></w:r></w:p></w:footnote>
```

Single para footnote rules:

w:p 

* attributes:
* w14:paraId="dddddddd" (8-digit alpha-numeric unique string) 
* w14:textId="77777777" (always 77777777)
* w:rsidR="dddddddd" (8-digit alpha-numeric string, shared with all w:p elements in footnotes section; starts with 00)
* w:rsidRDefault="dddddddd" (8-digit alpha-numeric string, shared with all w:p elements in footnotes section; SAME AS w:rsidR)

Run rules:

IF run begins with a space, first child must be `<w:t xml:space="preserve">`

Runs containing character styles are standalone, and must include: w:rsidRPr="dddddddd" (8-digit alpha-numeric unique string)

E.g.: `<w:r w:rsidRPr="00990792"><w:rPr><w:rStyle w:val="spanitaliccharactersital"/></w:rPr><w:t>someday I shall run my fingers through my lover’s hair</w:t></w:r>`

Multi para footnote rules:

Same attributes as single, PLUS:

* w:rsidRPr="006C2FBF" (8-digit alpha-numeric string, shared with all w:p elements in this specific footnote)
* w:rsidP="006C2FBF" (8-digit alpha-numeric string, shared with all w:p elements in this specific footnote; SAME AS rsidRPr)

EXCEPT last para in the multi-para footnote, which follows the single-para structure.

E.g.:

```
<w:footnote w:id="7">
  <w:p w14:paraId="60461657" w14:textId="77777777" w:rsidR="00A51F92" w:rsidRPr="006C2FBF" w:rsidRDefault="00A51F92" w:rsidP="006C2FBF"><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r w:rsidRPr="006C2FBF"><w:t>You will knock, and a sharp-eyed old man answer. “Yes?” He’ll look you over and see what you hold: a fist-size pouch, fat with coin. “What do you want?”</w:t></w:r></w:p>
  <w:p w14:paraId="7BD3F9F1" w14:textId="77777777" w:rsidR="00A51F92" w:rsidRPr="006C2FBF" w:rsidRDefault="00A51F92" w:rsidP="006C2FBF"><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r w:rsidRPr="006C2FBF"><w:t>“To talk to the lady of the warrior Cumalo, please. I was a friend of his.”</w:t></w:r></w:p>
  <w:p w14:paraId="210CF642" w14:textId="77777777" w:rsidR="00A51F92" w:rsidRDefault="00A51F92" w:rsidP="006C2FBF"><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="Times New Roman"/></w:rPr><w:t>“Ma’am . . . ,” you’ll say, holding out the little bag of full-weights. “Your husband . . .” Then your voice clots, and your eyes spill the first tears they will have shed for poor Cumalo; Janisse shaking her head slowly at first, and then with vehemence. These will be the gesture and instant comprehension of someone who’s always known she’d have to hear what’s about to be said. The child, too young to understand, sleepily sucks her thumb, looking between weeping mother, weeping stranger.</w:t></w:r></w:p>
</w:footnote>
```

IN document.xml, footnote ref is a standalone run:

```
<w:r w:rsidR="009A405C" w:rsidRPr="009A405C"><w:rPr><w:rStyle w:val="spansuperscriptcharacterssup"/></w:rPr><w:footnoteReference w:id="1"/></w:r>
```

Surrounding text are standard runs.

## DOCUMENT.XML.RELS

DOCUMENT settings:

All pieces of the package must be listed in the word/_rels/document.xml.rels file, in this order:

1. endnotes.xml
1. styles.xml
1. footnotes.xml
1. theme1.xml
1. numbering.xml
1. ../customXml/item1.xml
1. webSettings.xml
1. fontTable.xml
1. settings.xml
1. footer2.xml
1. stylesWithEffects.xml
1. footer1.xml

With this notation: `<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>`

Each of those items will be numbered sequentially, in this format: Id="rId1", Id="rId2", Id="rId3" ... etc.

## [CONTENT_TYPES].XML

Footnotes and endnotes must also be list in the [Content_Types].xml file, like so:

```
<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
```

These can be inserted anywhere within the `<Types>` parent element.