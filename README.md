# PP_LGManager
This is the language manager plugin for PowerPoint.

The PP_LGManager allows the use of multiple languages within one PowerPoint Document, by  defining an XML structure for all shapes.
Each Shape is stored as an element node inside the .xml file, whereas the text of the shape is stored as a text node. Each new language is stored as a seperate text field as well as a global language node at the start of the document..

## How To Use It
Go to File -> PowerPoint Options -> Add-Ins -> Select "PowerPoint-AddIns" in the Manage tab -> Add new -> select the .ppam file. 
If you are asked about Makro safety, confirm and set to an appropriate level. You will also have to allow the access to the VBA-Projektmodell.

After this you should see a new Ribbon tab called LG_Manager, including 3 buttons.

### Load
Load will check for an .xml file (with the same name as the .ppt) and look for language node. if no nodes are found or no .xml file is found it will tell you so.
Otherwise you can select a language and the script will change all text fields (no text in images of course) according to the stored text nodes.

### Save
If no .xml is found a new one will be generated. You will also have to set the first language.
Otherwise you can either select an already created language and overwrite all textfields for it, or create and save it as a new language.

### 
The Handler button toggles the state of the handler: If active, the normal Save action from PowerPoint (for example when pressing [CTRL]-[S]) will also trigger the script save language method. The same is applied when loading a document.
After closing PowerPoint the Handler will be deactivated again.
If the Handler is deactivated nothing happens on PP Save /Load.

### Format
Element NAME="Titel 1"
	Text Language="deutsch"
		Fragment format="" font="" color="" bullet=""
			TEXT

format="alignment(l|c|r)|fontsize(1-xy)|bolditalic|underlined|embossed|shadowed|subscript|superscript|"
font="fontname"
colorIndex="nr"
bullets="??"

Necessary to compare:
ParagraphFormat: Alignment (bullet)
Font: bold, color.RGB, Emboss, italic, name, shadow, size, subscript, superscript, underline


###Known Issues / TODO:
- Fix last block slide to left 
- bulleting
- leading/tailing newlines are deleted
=> insert special char for new lines? check options?
find with regex pos(\ )
- save xml inside .ppt											- not possible inside vba, only per shell script / exe
- look for new pages to add											
- on Plugin Load: VBIDE check
