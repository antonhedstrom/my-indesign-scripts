#target indesign;
//VIBS.jsx
//An InDesign JavaScript
/*
Copied from SortParagraph (sample script for Indesign)
*/


var isVIBS = false;

try{ main (); }
catch (e) { alert (e.message + " (" + e.line + ")"); }

function main(){
  //Make certain that user interaction (display of dialogs, etc.) is turned on.
  app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
  if (app.documents.length !== 0){
    if (app.selection.length === 1){
      switch (app.selection[0].constructor.name){
        case "Text":
        case "TextFrame":
        case "InsertionPoint":
        case "Character":
        case "Word":
        case "Line":
        case "TextStyleRange":
        case "Paragraph":
        case "TextColumn":
          startScript();
          break;
        default:
          alert("The selected object ("+app.selection[0].constructor.name+") is not a text object. Select some text and try again.");
          break;
      }
    }
    else{
      alert("Please select some text and try again.");
    }
  }
  else{
    alert("No documents are open. Please open a document, select some text, and try again.");
  }
}
function startScript(){
  var sel = app.selection[0];
      story = sel.parentStory;

  cleanShit(story);
  addTombstone(story);
  addAuthor(story);
  setParagraphStyles(story.paragraphs);
}

/*
  Clean up as much as possible.

  Flags:
   g = Global
   m = multiline
*/
function cleanShit(story) {
  var text = story.contents;

  var regexps = [
    {pattern: /\s+$/gm, replaceWith: ''}, // Trailing spaces
    {pattern: /^\s+/gm, replaceWith: ''}, // Leading spaces in paragraphs and duplicated newlines
    {pattern: /([ ]){2,}/gm, replaceWith: ' '}, // Duplicate spaces
   // {pattern: /(\r){2,}/gm, replaceWith: '\r'}, // Duplicate new lines
    //{pattern: /(~b~b+)/g, replaceWith: '\r'}, // Duplicated new lines (from default query in find/replace in Indesign)
  ];
  for ( var i in regexps) {
    text = text.replace(regexps[i].pattern, regexps[i].replaceWith);
  }

  story.contents = text;
  return;
}


function addTombstone(story) {
    var noneStyle = this.document.characterStyles.itemByName("[None]");
    var tombstoneStyle = getCharacterStyleByName("Tombstone");
    story.insertionPoints.lastItem().applyCharacterStyle(tombstoneStyle);
    story.insertionPoints.lastItem().contents = " ■";
    story.insertionPoints.lastItem().applyCharacterStyle(noneStyle);
}

function addAuthor(story) {
  var defaultArticleParagraphStyle = "Article:Body:ARTI Body main";
  var authorStyles = {
    name: "Article:Author:ARTI Body author name",
    email: "Article:Author:ARTI Body author email"
  };

  var author = getAuthorByPrompt();

  story.insertionPoints.lastItem().contents = "\r" + author.name;
  story.paragraphs.lastItem().applyParagraphStyle(getParagraphStyleByName(authorStyles.name));

  if ( author.email ) {
    story.insertionPoints.lastItem().contents = "\r" + author.email;
    story.insertionPoints.lastItem().applyParagraphStyle(getParagraphStyleByName(authorStyles.email));
  }
  if ( author.title ) {
    story.insertionPoints.lastItem().contents = "\r" + author.title;
    story.insertionPoints.lastItem().applyParagraphStyle(getParagraphStyleByName(authorStyles.email));
  }

}

function setParagraphStyles(paragraphs) {
  var pStyle,
      styleName,
      pStyles = {
        first: "Article:ARTI Stock Info",
        second: "Article:ARTI Body lead with drop char",
        header: "Article:ARTI Body header", // Unused?
        afterHeader: "Article:ARTI Body lead", // Unused?
        rest: "Article:ARTI Body main",
      };

  // If there is a VIBS or Markn. article: remove "stock info" style.
  if ( isVIBS ) {
    pStyles = {
      first: "Article:VIBS:ARTI VIBS Body lead",
      header: "Article:VIBS:ARTI Body header", // Unused?
      afterHeader: "Article:VIBS:ARTI Body lead", // Unused?
      rest: "Article:VIBS:ARTI VIBS Body main",
    };
  }

  for ( var i = 0; i < paragraphs.length; i++) {
    styleName = undefined;
    switch(i) {
      case 0:
        styleName = pStyles.first;
        break;
      case 1:
        styleName = pStyles.second || pStyles.rest;
        break;
      case paragraphs.length:
        styleName = pStyles.last || undefined;
        break;
      default:
        styleName = pStyles.rest;
    }

    if ( styleName ) {
      pStyle = getParagraphStyleByName(styleName);
      paragraphs[i].applyParagraphStyle(pStyle);
    }
  }


}


// ################## HELPER FUNCTIONS #################################

function getParagraphStyleByName(name) {
  var paragraphs,
      currentGroup = this.document,
      psName,
      elements = name.split(":"),
      groups = [];
  if ( elements.length === 1 ) {
    psName = name;
  }
  else {
    psName = elements.slice(-1)[0]; // Get the last (name)
    groups = elements.slice(0, -1); // All but the last (all groups)
  }


  try {
    if ( groups.length === 0 ) {
      paragraphs = this.document.paragrapghStyles;
    }
    else {
      for (var i in groups) {
        currentGroup = currentGroup.paragraphStyleGroups.item(groups[i]);
      }
      paragraphs = currentGroup.paragraphStyles;
    }
    return paragraphs.item(psName);
  } catch (err) {
    alert(err);
  }

  alert("Paragraph name ("+name+") not found!");
}



function getCharacterStyleByName(name) {
  var characterStyles,
      characterStyleName,
      currentGroup = this.document,
      elements = name.split(":"),
      groups = [];
  if ( elements.length === 1 ) {
    characterStyleName = name;
  }
  else {
    characterStyleName = elements.slice(-1)[0]; // Get the last (name)
    groups = elements.slice(0, -1); // All but the last (all groups)
  }


  try {
    if ( groups.length === 0 ) {
      characterStyles = this.document.characterStyles;
    }
    else {
      for (var i in groups) {
        currentGroup = currentGroup.characterStyleGroups.item(groups[i]);
      }
      characterStyles = currentGroup.characterStyles;
    }
    return characterStyles.item(characterStyleName);
  } catch (err) {
    alert(err);
  }

  alert("Character name ("+name+") not found!");
}



function getAuthorByPrompt() {
   var authors = {
    "jacek": {
      name: "Jacek Bielecki",
      email: "jacek@stockpicker.se"
    },
    "fredrik": {
      name: "Fredrik Aabye-Nielsen",
      email: "fredrik@stockpicker.se"
    },
    "jan": {
      name: "Jan Axelsson",
      email: null
    },
    "per": {
      name: "Per Bernhult",
      email: null
    },
    "kim": {
      name: "Kim Dahlström",
      email: null
    }
  };

  var w = new Window ('dialog', "Ange författare", undefined, {close: true});
  w.alignChildren = "right"; w.maximumSize.height = 550, w.minimumSize.width = 200;

  var buttons = w.add("group");
  var result;


  for (var i in authors) {
    buttons.add('button', undefined, authors[i].name || i, {name: 'author-'+i}).onClick = function(clickedAuthor) {
       return function() {
           result = clickedAuthor;
           exit();
           w.close();
       };
    }(authors[i]);
  }

  // var controllers = w.add("group");
  //controllers.add('checkbox');

  w.show();
  return result;

}
