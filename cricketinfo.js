// node cricketinfo.js --source=https://www.espncricinfo.com/series/icc-cricket-world-cup-2019-1144415/match-results --excel=WC19.xlsx --datafolder=fixture

// node i minimist
// node i excel4node
// node i axios
// node i pdf-lib
let minimist =require("minimist");
let axios = require("axios"); //download
let jsdom = require("jsdom"); // read
let excel4node = require("excel4node"); // excel creation
let  pdf = require ("pdf-lib"); // pdf create
let fs = require("fs");
let path = require("path");

let args=minimist(process.argv);

let whenfulfilled= axios.get(args.source);
whenfulfilled.then(function(response){
    let url=response.data;
    
    let dom = new jsdom.JSDOM(url);
    let document = dom.window.document;
    
    let matches = [];
    let matchescores = document.querySelectorAll("div.match-score-block");
    for(let i=0; i<matchescores.length;i++){
        let match={
        };

        let name = matchescores[i].querySelectorAll(" p.name");
         match.team1=name[0].textContent;   
         match.team2=name[1].textContent;   

         let score= matchescores[i].querySelectorAll("div.score-detail > span.score ");
         if (score.length==2){
            match.team1s=score[0].textContent;   
            match.team2s=score[1].textContent;
         }else if(score.length==1){
            match.team1s=score[0].textContent;   
            match.team2s="Did not bat";
         }else{
            match.team1s="NO PLAY";   
            match.team2s="NO PLAY";
         }
         let result=matchescores[i].querySelector("div.status-text > span");
         match.result = result.textContent;

            matches.push(match);
         
            
      }
      let matchesJSON=JSON.stringify(matches);
      fs.writeFileSync("matches.json", matchesJSON, "utf-8");

        let teams= [];
        for(let i=0; i < matches.length; i++){
          puttingTeams(teams,matches[i]);
   }
   for(let i=0; i < matches.length; i++){
      puttingmatches(teams,matches[i]);
   }
      let teamsJSON=JSON.stringify(teams);
      fs.writeFileSync("teams.json", teamsJSON, "utf-8");
      createexcel(teams);
      createFolders(teams);
   })
function createFolders(teams,dataDir){
   if(fs.existsSync(dataDir)==true){
      fs.rmdirSync(dataDir,{recursive : true});
   }
   fs.mkdirSync(args.datafolder);

    for(i=0;i<teams.length;i++){
     let Teamfolder=path.join(args.datafolder,teams[i].name);
     fs.mkdirSync(Teamfolder);

    for(let j=0;j< teams[i].matches.length;j++){
         let matchfile = path.join(Teamfolder, teams[i].matches[j].vs+ ".pdf");
         createscorecard(teams[i].name, teams[i].matches[j],matchfile);
 }  
    }

}

function createscorecard(teamname,match,matchfile){
   let  t1= teamname;
   let t2=match.vs;
   let t1s= match.self;
   let t2s= match.opponent;
   let result= match.result;

   let bytesPDFtemplate= fs.readFileSync("template.pdf");
   let pdfpromise=pdf.PDFDocument.load(bytesPDFtemplate);
   pdfpromise.then(function(pdfdoc){
      let page= pdfdoc.getPage(0);
           
      page.drawText(t1,{
         x:380,
         y:655,
         size: 20
      });
      page.drawText(t2,{
         x:380,
         y:619,
         size: 20
      });
      page.drawText(t1s,{
         x:380,
         y:589,
         size: 20
      });
      page.drawText(t2s,{
         x:380,
         y:558,
         size: 20
      });
      page.drawText(result,{
         x:340,
         y: 532,
         size: 20
      });
      let finalpdf= pdfdoc.save();
      finalpdf.then(function(finalPDFbytes){
         if(fs.existsSync(matchfile + ".pdf")== true){
            fs.writeFileSync(matchfile + "1.pdf", finalPDFbytes);
         }else{
         fs.writeFileSync(matchfile + ".pdf", finalPDFbytes);
         }
      })
   })
}

   function createexcel(teams){
      let wb=new excel4node.Workbook();
      for(let i=0;i< teams.length;i++){
         let sheet = wb.addWorksheet(teams[i].name);

         sheet.cell(1,1).string("VS");
         sheet.cell(1,2).string("Self score");
         sheet.cell(1,3).string("Opp score");
         sheet.cell(1,4).string("Result");
          for(let j=0;j< teams[i].matches.length;j++){
            sheet.cell(2+j,1).string(teams[i].matches[j].vs);
            sheet.cell(2+j,2).string(teams[i].matches[j].self);
            sheet.cell(2+j,3).string(teams[i].matches[j].opponent);
            sheet.cell(2+j,4).string(teams[i].matches[j].result);
          }
      }
      wb.write(args.excel);
   }

function puttingTeams(teams,match){
   let t1index= -1;
   for(let i=0;i< teams.length;i++){
      if(teams[i].name==match.team1){
         t1index=i;
         break;
      }
   }
   if(t1index==-1){
      let team={
         name: match.team1,
         matches: []
      };
      teams.push(team);
   }
   let t2index= -1;
   for(let i=0;i< teams.length;i++){
      if(teams[i].name==match.team2){
         t2index=i;
         break;
      }
   }

   if(t2index==-1){
      let team={
         name: match.team2,
         matches: []
      };
      teams.push(team);
   }

}
function puttingmatches(teams,match){
      let t1index=-1;
      for(let i=0;i< teams.length;i++){
         if(teams[i].name==match.team1){
            t1index=i;
            break;
         }
      }

    let team1 = teams[t1index];
    team1.matches.push({
       vs: match.team2,
       self: match.team1s,
       opponent: match.team2s,
       result: match.result
    });
    let t2index=-1;
    for(let i=0; i < teams.length; i++){
      if(teams[i].name==match.team2){
         t2index=i;
         break;
      }
   } 
    let team2 = teams[t2index];
    team2.matches.push({
       vs: match.team1,
       self: match.team2s,
       opponent: match.team1s,
       result: match.result
    });
}
