#target indesign;
//VIBS.jsx
//An InDesign JavaScript
/*
Copied from SortParagraph (sample script for Indesign)
*/

// Fix a number of things related to an VIBS article:
// * Set first paragraph to use style "ARTI VIBS Body lead"
// * Add tombstone at end of article
// * Add Jacek as author
//

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

  cleanShit();
  addTombstone(story);
  addAuthor(story);
  setParagraphStyles(story.paragraphs);
}

/*
  Clean up as much as possible. Ie:
  * Duplicated "new paragraphs characters
  * New line at end of text
*/
function cleanShit() {
  // Todo
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
    name: "Article:Body:ARTI Body author name",
    email: "Article:Body:ARTI Body author email"
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
  var paragraphStylesToApply = [
    "Article:ARTI Stock Info", // First element will be applied to first paragrapgh
    "Article:ARTI VIBS Body lead", // Second element will be applied to second paragrapgh
  ];
  var pg;
  var pgStyle;

  for ( var i = 0; i < paragraphs.length; i++) {


  for ( var i = 0; i < paragraphStylesToApply.length; i++) {
    pg = paragraphs[i];
    pgStyle =  getParagraphStyleByName(paragraphStylesToApply[i]);
    pg.applyParagraphStyle(pgStyle);
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

  w.show();
  return result;

}
