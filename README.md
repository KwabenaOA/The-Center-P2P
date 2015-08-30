# The-Center-P2P
//Peer to Student Matching Program for The Center at Oakton High School
//Senior Legacy Project
//The Center Class 2014-15

var tutors = new Array();
var tutees = new Array();
var tutorsMentors = new Array();
var tutorsTutors = new Array();
var tuteesMentors = new Array();
var tuteesTutors = new Array();

var spreadGet = new Array();
var spreadSet = new Array();
spreadSet.push(["Tutor Name","","Tutee Name","Matched Previously or Now?","Subjects in Common","Subjects in Contrast","Extracurriculars in Common"]);
spreadSet.push(["","","","","","",""]);

//doc1 is for the tutors
var doc1 = SpreadsheetApp.openById("14JdHZtU3fSEkm5UdKpLAVQiro6wpT-uoP9Xj0fe2SPk");
//doc2 is for the tutees
var doc2 = SpreadsheetApp.openById("1i1svftw72FHiIC8aQM4COP2cAFQ_FREkRUc7bcNU7uQ");
//doc3 will be the document for the matched tutors and tutees
var doc3 = new Array();

//The tutor class
function tutor(name, status, grade, bestSubs, worstSubs, extraCur, isMatched, matchedName)
{
  this.name = name;
  this.status = status;
  this.grade = grade;
  this.bestSubs = bestSubs;
  this.worstSubs = worstSubs;
  this.extraCur = extraCur;
  this.isMatched = isMatched;
  this.matchedName = matchedName;
  
  var bestSubsArray = separateSubs(this.bestSubs);
  var worstSubsArray = separateSubs(this.worstSubs);
  var extraCurArray = separateSubs(this.extraCur);
  
  this.bestSubsArray = bestSubsArray;
  this.worstSubsArray = worstSubsArray;
  this.extraCurArray = extraCurArray
}
tutor.prototype.getBestSubs = function() {return this.bestSubs};
tutor.prototype.getWorstSubs = function() {return this.worstSubs};
tutor.prototype.getExtraCur = function() {return this.extraCur};
tutor.prototype.getName = function() {return this.name};
tutor.prototype.getStatus = function() {return this.status};
tutor.prototype.getGrade = function() {return this.grade};
tutor.prototype.getBestSubsArray = function() {return this.bestSubsArray};
tutor.prototype.getWorstSubsArray = function() {return this.worstSubsArray};
tutor.prototype.getExtraCurArray = function() {return this.extraCurArray};
tutor.prototype.getIsMatched = function() {return this.isMatched};
tutor.prototype.getMatchedName = function() {return this.matchedName};
//The tutee class
function tutee(name, status, grade, allSubs, extraCur)
{ 
  this.name = name;
  this.status = status;
  this.grade = grade;
  this.allSubs = allSubs;
  this.extraCur = extraCur;
  
  var allSubsArray = separateSubs(allSubs);
  var extraCurArray = separateSubs(this.extraCur);
  
  this.allSubsArray = allSubsArray;
  this.extraCurArray = extraCurArray
}
tutee.prototype.getAllSubs = function() {return this.allSubs};
tutee.prototype.getExtraCur = function() {return this.extraCur};
tutee.prototype.getName = function() {return this.name};
tutee.prototype.getStatus = function() {return this.status};
tutee.prototype.getGrade = function() {return this.grade};
tutee.prototype.getAllSubsArray = function() {return this.allSubsArray};
tutee.prototype.getExtraCurArray = function() {return this.extraCurArray};
//Creating the doc3 spreadsheet
function createDoc()
{
  var monthNames = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  var date = new Date();
  var day = date.getDate();
  var monthIndex = date.getMonth();
  var year = date.getFullYear();
  
  doc3 = SpreadsheetApp.create("P2P Match (" + monthNames[monthIndex] + " " + day + ", " + year + ")");
}
//This separates the long strings of text when selecting subjects
function separateSubs(choices)
{
  var j = 0;
  var commaIndex = 0;
  var preCommaIndex = 0;
  var allChoices = new Array();
  while(true)
  {
    preCommaIndex = commaIndex;
    if(j == 0)
      commaIndex = choices.indexOf(",", preCommaIndex);
    else
      commaIndex = choices.indexOf(",", (preCommaIndex+1));
    if(commaIndex == -1)
      break;
    if(j == 0)
      allChoices.push(choices.substring(preCommaIndex, commaIndex));
    else
      allChoices.push(choices.substring((preCommaIndex+2), commaIndex));
    j++;
  }
  if(j == 0)
    allChoices.push(choices.substring(preCommaIndex));
  else
    allChoices.push(choices.substring(preCommaIndex+2));
  return allChoices;
}
//Creates an array of the tutors
function getTutors()
{
  spreadGet = doc1.getDataRange().getValues();
  for(var k = 1; k < spreadGet.length; k++)
  {
    if(spreadGet[k][7].equals("No"))
      tutors.push(new tutor(spreadGet[k][1],spreadGet[k][4],spreadGet[k][6],spreadGet[k][2],spreadGet[k][3],spreadGet[k][5],spreadGet[k][7],""));
    else
      tutors.push(new tutor(spreadGet[k][1],spreadGet[k][4],spreadGet[k][6],spreadGet[k][2],spreadGet[k][3],spreadGet[k][5],spreadGet[k][7],spreadGet[k][8]));
  }
}
//Creates an array of the tutees
function getTutees()
{
  spreadGet = doc2.getDataRange().getValues();
  for(var k = 1; k < spreadGet.length; k++)
  {
    if(spreadGet[k][6].equals("Mentor"))
      tutees.push(new tutee(spreadGet[k][1],spreadGet[k][6],spreadGet[k][2],"",spreadGet[k][7]));
    else if(spreadGet[k][6].equals("Tutor"))
      tutees.push(new tutee(spreadGet[k][1],spreadGet[k][6],spreadGet[k][2],spreadGet[k][3],""));
    else if(spreadGet[k][6].equals("Both"))
      tutees.push(new tutee(spreadGet[k][1],spreadGet[k][6],spreadGet[k][2],spreadGet[k][9],spreadGet[k][8]));
  }
}
//Provides the names of the subjects in common
function subsCompare(tutor, tutee)
{
  var string = "";
  var compare = new Array();
  for(var k = 0; k < tutor.getBestSubsArray().length; k++)
  {
    for(var j = 0; j < tutee.getAllSubsArray().length; j++)
    {
      if(tutor.getBestSubsArray()[k].equals(tutee.getAllSubsArray()[j]))
        compare.push(tutor.getBestSubsArray()[k]);
    }
  }
  if(compare.length > 0)
  {
    for(var k = 0; k < compare.length; k++)
    {
      if(k != (compare.length-1))
        string += compare[k] + ", ";
      else
        string += compare[k];
    }
    return string;
  }
  else
    return "No agreeing subjects"
}
//Provides the names of the subjects which aren't in common
function subsContrast(tutor, tutee)
{
  var string = "";
  var contrast = new Array();
  for(var k = 0; k < tutor.getWorstSubsArray().length; k++)
  {
    for(var j = 0; j < tutee.getAllSubsArray().length; j++)
    {
      if(tutor.getWorstSubsArray()[k].equals(tutee.getAllSubsArray()[j]))
        contrast.push(tutor.getWorstSubsArray()[k]);
    }
  }
  if(contrast.length > 0)    
  {
    for(var k = 0; k < contrast.length; k++)
    {
      if(k != (contrast.length-1))
        string += contrast[k] + ", ";
      else
        string += contrast[k];
    }
    return string;
  }
  else
    return "No contrasting subjects"
    }
function extrasCompare(tutor, tutee)
{
  var string = "";
  var compare = new Array();
  for(var k = 0; k < tutor.getExtraCurArray().length; k++)
  {
    for(var j = 0; j < tutee.getExtraCurArray().length; j++)
    {
      if(tutor.getExtraCurArray()[k].equals(tutee.getExtraCurArray()[j]))
        compare.push(tutor.getExtraCurArray()[k]);
    }
  }
  if(compare.length > 0)
  {
    for(var k = 0; k < compare.length; k++)
    {
      if(k != (compare.length-1))
        string += compare[k] + ", ";
      else
        string += compare[k];
    }
    return string;
  }
  else
    return "No agreeing extracurriculars"
}
//Matches tutors with tutees if they were already matched before
function prevMatch()
{
  var k = 0;
  
  while(k < tutors.length)
  {
    if(tutors[k].getIsMatched().equals("Yes"))
    {
      for(var j = 0; j < tutees.length; j++)
      {
        if(tutors[k].getMatchedName().equals(tutees[j].getName()))
        {
          var centerName = tutors[k].getName();
          if(tutors[k].getStatus().equals("Mentor"))
            centerName += " (M)";
          else if(tutors[k].getStatus().equals("Tutor"))
            centerName += " (T)";
          else if(tutors[k].getStatus().equals("Both"))
            centerName += " (B)";
        
          if(tutees[j].getStatus().equals("Mentor"))
            spreadSet.push([centerName,"<= Matched =>",tutees[j].getName() + " (M)","Previously","N/A","N/A",extrasCompare(tutors[k],tutees[j])]);
          else if(tutees[j].getStatus().equals("Tutor"))
            spreadSet.push([centerName,"<= Matched =>",tutees[j].getName() + " (T)","Previously",subsCompare(tutors[k],tutees[j]),subsContrast(tutors[k],tutees[j]),"N/A",]);
          else if(tutees[j].getStatus().equals("Both"))
            spreadSet.push([centerName,"<= Matched =>",tutees[j].getName() + " (B)","Previously",subsCompare(tutors[k],tutees[j]),subsContrast(tutors[k],tutees[j]),extrasCompare(tutors[k],tutees[j])]);
          tutors.splice(k,1);
          tutees.splice(j,1);
          k--;
          break;
        }
      }
    }
    k++;
  }
}
//Creates an array of percentages of tutor relationships with respect to one tutee
function ranksSubs(tutorArray, tutee)
{
  var rank = 0;
  var percentage = 0.0;
  var tutor = new Array();
  var percents = new Array();
  var a = 0;
  while(a < tutorArray.length)
  {
    tutor = tutorArray[a];
    for(var k = 0; k < tutor.getBestSubsArray().length; k++)
    {
      for(var j = 0; j < tutee.getAllSubsArray().length; j++)
      {
        if(tutor.getBestSubsArray()[k].equals(tutee.getAllSubsArray()[j]))
          rank++;
      }
    }
    for(var k = 0; k < tutor.getWorstSubsArray().length; k++)
    {
      for(var j = 0; j < tutee.getAllSubsArray().length; j++)
      {
        if(tutor.getWorstSubsArray()[k].equals(tutee.getAllSubsArray()[j]))
          rank--;
      }
    }
    percentage = rank/(tutee.getAllSubsArray().length);
    percents.push(percentage);
    a++;
  }
  return percents;
}
//Creates an array of percentages of tutor relationships with respect to one tutee
function ranksExtraCur(tutorArray, tutee)
{
  var rank = 0;
  var percentage = 0.0;
  var tutor = new Array();
  var percents = new Array();
  var a = 0;
  while(a < tutorArray.length)
  {
    tutor = tutorArray[a];
    for(var k = 0; k < tutor.getExtraCur().length; k++)
    {
      for(var j = 0; j < tutee.getExtraCur().length; j++)
      {
        if(tutor.getExtraCur()[k].equals(tutee.getExtraCur()[j]))
          rank++;
      }
    }
    percentage = rank/(tutee.getExtraCur().length);
    percents.push(percentage);
    a++;
  }
  return percents;
}
//Creates an array of the arrays of indexes for the "best tutors" for each tutee
function bestTutors(tutorArray, tuteeArray)
{
  var Ranks1 = new Array();
  var Ranks2 = new Array();
  var BestTutors = new Array();
  var arrayOfBest = new Array();
  for(var k = 0; k < tuteeArray.length; k++)
  {
    if(tuteeArray[k].getStatus().equals("Tutor"))
    {
      Ranks1 = ranksSubs(tutorArray,tuteeArray[k]);
      var maxPercent = Ranks1[0];
      for(var j = 0; j < Ranks1.length; j++)
      {
        BestTutors.push(j);
        if(Ranks1[j] > maxPercent)
        {
          BestTutors = new Array();
          BestTutors.push(j);
          maxPercent = Ranks1[j];
        }
      }
    }
    else if(tuteeArray[k].getStatus().equals("Mentor"))
    {
      Ranks2 = ranksExtraCur(tutorArray,tuteeArray[k]);
      var maxPercent = Ranks2[0];
      for(var j = 0; j < Ranks2.length; j++)
      {
        BestTutors.push(j);
        if(Ranks2[j] > maxPercent)
        {
          BestTutors = new Array();
          BestTutors.push(j);
          maxPercent = Ranks2[j];
        }
      }
    }
    else if(tuteeArray[k].getStatus().equals("Both"))
    {
      Ranks1 = ranksSubs(tutorArray,tuteeArray[k]);
      Ranks2 = ranksExtraCur(tutorArray,tuteeArray[k]);
      var maxPercent = (Ranks1[0] + Ranks2[0]);
      for(var j = 0; j < Ranks1.length; j++)
      {
        BestTutors.push(j);
        if((Ranks1[j] + Ranks2[j]) > maxPercent)
        {
          BestTutors = new Array();
          BestTutors.push(j);
          maxPercent = (Ranks1[j] + Ranks2[j]);
        }
      }
    }
    arrayOfBest.push(BestTutors);
  }
  return arrayOfBest;
}
//Matches tutees with tutors
function match()
{
  var arrayOfTutees = new Array();
  var arrayOfTutors = new Array();
  var indexes = bestTutors(tutors,tutees);
  while(tutees.length > 0 && tutors.length > 0)
  {
    for(var k = 0; k < tutors.length; k++)
    {
      for(var i = 0; i < indexes.length; i++)
      {
        for(var j = 0; j < indexes[i].length; j++)
        {
          if(indexes[i][j] == k)
          {
            arrayOfTutees.push(tutees[i]);
            break;
          }
        }
      }
      if(arrayOfTutees.length > 0)
      {
        var rnd = Math.floor(Math.random()*(arrayOfTutees.length-1));
        
        var centerName = tutors[k].getName();
        var tuteeName = arrayOfTutees[rnd].getName();
        
        if(tutors[k].getStatus().equals("Mentor"))
          centerName += " (M)";
        else if(tutors[k].getStatus().equals("Tutor"))
          centerName += " (T)";
        else if(tutors[k].getStatus().equals("Both"))
          centerName += " (B)";
        
        if(arrayOfTutees[rnd].getStatus().equals("Mentor"))
          tuteeName += " (M)";
        else if(arrayOfTutees[rnd].getStatus().equals("Tutor"))
          tuteeName += " (T)";
        else if(arrayOfTutees[rnd].getStatus().equals("Both"))
          tuteeName += " (B)";
        
        if(arrayOfTutees[rnd].getStatus().equals("Mentor"))
          spreadSet.push([centerName,"<= Matched =>",tuteeName,"Now","N/A","N/A",extrasCompare(tutors[k],arrayOfTutees[rnd])]);
        else if(arrayOfTutees[rnd].getStatus().equals("Tutor"))
          spreadSet.push([centerName,"<= Matched =>",tuteeName,"Now",subsCompare(tutors[k],arrayOfTutees[rnd]),subsContrast(tutors[k],arrayOfTutees[rnd]),"N/A"]);
        else if(arrayOfTutees[rnd].getStatus().equals("Both"))
          spreadSet.push([centerName,"<= Matched =>",tuteeName,"Now",subsCompare(tutors[k],arrayOfTutees[rnd]),subsContrast(tutors[k],arrayOfTutees[rnd]),extrasCompare(tutors[k],arrayOfTutees[rnd])]);
        arrayOfTutors.push(tutors[k]);
        for(var i = 0; i < tutees.length; i++)
        {
          if(tutees[i].getName().equals(arrayOfTutees[rnd].getName()))
          {
            tutees.splice(i,1);
            indexes.splice(i,1);
              break;
          }
        }
        for(var p = 0; p < arrayOfTutors.length; p++)
        {
          for(var j = 0; j < tutors.length; j++)
          {
            if(arrayOfTutors[p].getName().equals(tutors[j].getName()))
            {
              tutors.splice(j,1);
              break;
            }
          }
        }
        arrayOfTutees = new Array();
      }
    }
    if(tutees.length == 0 || tutors.length == 0)
    {
      break;
    }
    indexes = bestTutors(tutors, tutees);
  }
}
function extras()
{
  var k = 0;
  if(tutors.length > 0)
  {
    spreadSet[0].push("","Unmatched Tutors:","","","","","","");
    spreadSet[1].push("","Name","Mentor or Tutor","Best Subjects","Worst Subjects","Extracurriculars","","");
    for(k = 0; k < tutors.length; k++)
    {
      var centerName = tutors[k].getName();
      if(tutors[k].getStatus().equals("Mentor"))
        centerName += " (M)";
      else if(tutors[k].getStatus().equals("Tutor"))
        centerName += " (T)";
      else if(tutors[k].getStatus().equals("Both"))
        centerName += " (B)";
      if(spreadSet[(k+2)].length > 0)
        spreadSet[(k+2)].push("",centerName,tutors[k].getStatus(),tutors[k].getBestSubs(),tutors[k].getWorstSubs(),tutors[k].getExtraCur(),"","");
      else
      {
        spreadSet.push(["","","","","","","",""]);
        spreadSet[(k+2)].push("",tutors[k].getName(),tutors[k].getStatus(),tutors[k].getBestSubs(),tutors[k].getWorstSubs(),tutors[k].getExtraCur(),"","");
      }
    }
  }
  else if(tutees.length > 0)
  {
    spreadSet[0].push("","Unmatched Tutees:","","","","","","");
    spreadSet[1].push("","Name","Wants a Mentor or Tutor","Subjects Needing Help In","Extracurriculars","","","");
    for(k = 0; k < tutees.length; k++)
    {
      var tuteeName = tutees[k].getName();
      if(tutees[k].getStatus().equals("Mentor"))
        tuteeName += " (M)";
      else if(tutees[k].getStatus().equals("Tutor"))
        tuteeName += " (T)";
      else if(tutees[k].getStatus().equals("Both"))
        tuteeName += " (B)";
      if(spreadSet[(k+2)].length > 0)
      {
        if(tutees[k].getStatus().equals("Mentor"))
          spreadSet[(k+2)].push("",tuteeName,tutees[k].getStatus(),"N/A",tutees[k].getExtraCur(),"","","");
        else if(tutees[k].getStatus().equals("Tutor"))
          spreadSet[(k+2)].push("",tuteeName,tutees[k].getStatus(),tutees[k].getAllSubs(),"N/A","","","");
        else if(tutees[k].getStatus().equals("Both"))
          spreadSet[(k+2)].push("",tuteeName,tutees[k].getStatus(),tutees[k].getAllSubs(),tutees[k].getExtraCur(),"","","");
      }
      else
      {
        spreadSet.push(["","","","","","","",""]);
        if(tutees[k].getStatus().equals("Mentor"))
          spreadSet[(k+2)].push("",tuteeName,tutees[k].getStatus(),"N/A",tutees[k].getExtraCur(),"","","");
        else if(tutees[k].getStatus().equals("Tutor"))
          spreadSet[(k+2)].push("",tuteeName,tutees[k].getStatus(),tutees[k].getAllSubs(),"N/A","","","");
        else if(tutees[k].getStatus().equals("Both"))
          spreadSet[(k+2)].push("",tuteeName,tutees[k].getStatus(),tutees[k].getAllSubs(),tutees[k].getExtraCur(),"","","");
      }
    }
  }
  while((k+2) < spreadSet.length)
  {
    spreadSet[(k+2)].push("","","","","","","","");
    k++;
  }
}
//Sets up the doc3 spreadsheet
function setSpread()
{
  var p = 0;
  var alignments = new Array();
  var bold = new Array();
  var italics = new Array();
  
  for(var k = 0; k < spreadSet.length; k++)
  {
    alignments.push(["center","center","center","center","center","center","center"]);
  }
    
  bold.push(["bold","bold","bold","bold","bold","bold","bold"]);
  bold.push(["normal","normal","normal","normal","normal","normal","normal"]);
  for(var k = 2; k < spreadSet.length; k++)
  {
    if(spreadSet[k][0].equals(""))
      break;
    bold.push(["normal","bold","normal","normal","normal","normal","normal"]);
  }
  
  italics.push(["normal","normal","normal","normal","normal","normal","normal"]);
  italics.push(["normal","normal","normal","normal","normal","normal","normal"]);
  for(var k = 2; k < spreadSet.length; k++)
  {
    if(spreadSet[k][0].equals(""))
      break;
    italics.push(["italic","normal","italic","normal","normal","normal","normal"]);
  }
  
  if(tutors.length > 0 || tutees.length > 0)
  {
    alignments[0].push("center","center","center","center","center","center","center","center");
    alignments[1].push("center","center","center","center","center","center","center","center");
    
    bold[0].push("normal","normal","normal","normal","normal","normal","normal","normal");
    bold[1].push("normal","bold","bold","bold","bold","bold","normal","normal");
    
    italics[0].push("normal","normal","normal","normal","normal","normal","normal","normal");
    italics[1].push("normal","normal","normal","normal","normal","normal","normal","normal");
    
    for(p = 2; p < spreadSet.length; p++)
    {
        alignments[p].push("center","center","center","center","center","center","center","center");
        bold[p].push("normal","normal","normal","normal","normal","normal","normal","normal");
        italics[p].push("normal","italic","normal","normal","normal","normal","normal","normal");
    }
  }
  doc3.getActiveRange().setHorizontalAlignments(alignments);
  doc3.getActiveRange().setFontWeights(bold);
  doc3.getActiveRange().setFontStyles(italics);
  doc3.getActiveRange().getCell(1, spreadSet[0].length - (6)).setFontLine("underline");
  
  for(var k = 1; k <= spreadSet[0].length; k++)
  {
    doc3.autoResizeColumn(k);
  }
}
function myFunction()
{
  createDoc();
  getTutors();
  getTutees();
  prevMatch();
  
  match();
  
  extras();
  
  doc3.setActiveRange(doc3.getRange("R1C1:R" + spreadSet.length + "C15"));
  doc3.getActiveRange().setValues(spreadSet);
  
  setSpread();
  
  DriveApp.getFolderById("0B8Ce2qN08R_5a1lpdVZGVU10OFk").addFile(DriveApp.getFileById(doc3.getId()));
  DriveApp.getRootFolder().removeFile(DriveApp.getFileById(doc3.getId()));
}
