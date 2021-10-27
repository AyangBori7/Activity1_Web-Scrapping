//The purpose of this Project is to extract information of WorldCup 2019 from Cricinfo website and present that in the form of excel
//and pdf scorecards.
//The real purpose is to learn how to extract information and get experience with Javascript
//A very good reason to ever make a project is to have good fun


//npm init -y
//npm install minimist
//npm install axios
//npm install jsdom
//npm install excel4node
//npm install pdf-lib

//node Cric_Info_Extractor.js --excel=Worldcup.csv --dataFolder=data --source=https://www.espncricinfo.com/series/icc-cricket-world-cup-2019-1144415/match-results 

let minimist = require("minimist");
let axios = require("axios");
let jsdom = require("jsdom");
let excel = require("excel4node");
let pdf = require("pdf-lib");
let fs = require("fs");
let path = require("path");

let args = minimist(process.argv);

//download using axios
//read using jsdom
//make excel using excel4node
//make pdf using pdf-lib


let responsePromise = axios.get(args.source);

responsePromise.then(function (response) {
    let html = response.data;

    let dom = new jsdom.JSDOM(html);
    let document = dom.window.document;

    let matches = [];
    let matchScoreDivs = document.querySelectorAll("div.match-score-block");
    for (let i = 0; i < matchScoreDivs.length; i++) {
        let match = {
        };

        let nameps = matchScoreDivs[i].querySelectorAll("p.name");
        match.t1 = nameps[0].textContent;
        match.t2 = nameps[1].textContent;

        let scoreSpans = matchScoreDivs[i].querySelectorAll("span.score");
        if (scoreSpans.length == 2) {
            match.t1s = scoreSpans[0].textContent;
            match.t2s = scoreSpans[1].textContent;

        }
        else if (scoreSpans.length == 1) {
            match.t1s = scoreSpans[0].textContent;
            match.t2s = "";

        }
        else {
            match.t1s = "";
            match.t2s = "";
        }

        let spanResult = matchScoreDivs[i].querySelector("div.status-text > span")
        match.result = spanResult.textContent;

        matches.push(match);

    }


    let teams = [];
    for (let i = 0; i < matchScoreDivs.length; i++) {
        populateTeams(teams, matches[i]);
    }

    for (let i = 0; i < matchScoreDivs.length; i++) {
        populateMatchesinTeams(teams, matches[i]);
    }

    createExcelFile(teams);
    createFolders(teams);

}).catch(function (err) {
    console.log(err);
})

function createFolders(teams) {
    fs.mkdirSync(args.dataFolder);
    for (let i = 0; i < teams.length; i++) {
        let teamFN = path.join(args.dataFolder, teams[i].name);
        fs.mkdirSync(teamFN);

        for (let j = 0; j < teams[i].matches.length; j++) {
            let matchFileName = path.join(teamFN, teams[i].matches[j].vs);
            createScoreCard(teams[i].name, teams[i].matches[j], matchFileName);
        }
    }

}

function createScoreCard(teamName, match, matchFileName) {
    let t1 = teamName;
    let t2 = match.vs;
    let t1s = match.selfScore;
    let t2s = match.oppScore;
    let result = match.result;

    let bytesofPDFTemplate = fs.readFileSync("Template.pdf");
    let pdfdocPromise = pdf.PDFDocument.load(bytesofPDFTemplate);
    pdfdocPromise.then(function (pdfdoc) {
        let page = pdfdoc.getPage(0);

        page.drawText(t1, {
            x: 320,
            y: 670,
            size: 10
        });
        page.drawText(t2, {
            x: 320,
            y: 640,
            size: 10
        });
        page.drawText(t1s, {
            x: 320,
            y: 610,
            size: 10
        });
        page.drawText(t2s, {
            x: 320,
            y: 580,
            size: 10
        });
        page.drawText(result, {
            x: 320,
            y: 550,
            size: 10
        });

        let finalPDFBytesPromise = pdfdoc.save();
        finalPDFBytesPromise.then(function (finalPDFBytes) {
            if(fs.existsSync(matchFileName + ".pdf") == true){
                fs.writeFileSync(matchFileName + "1.pdf", finalPDFBytes);
            } else {
                fs.writeFileSync(matchFileName + ".pdf", finalPDFBytes);
            }
        })


    })
}

function createExcelFile(teams) {
    let wb = new excel.Workbook();

    for (let i = 0; i < teams.length; i++) {
        let sheet = wb.addWorksheet(teams[i].name);
        sheet.cell(1, 1).string("VS");
        sheet.cell(1, 2).string("Self Score");
        sheet.cell(1, 3).string("Opp Score");
        sheet.cell(1, 4).string("Result");

        for (let j = 0; j < teams[i].matches.length; j++) {
            sheet.cell(2 + j, 1).string(teams[i].matches[j].vs);
            sheet.cell(2 + j, 2).string(teams[i].matches[j].selfScore);
            sheet.cell(2 + j, 3).string(teams[i].matches[j].oppScore);
            sheet.cell(2 + j, 4).string(teams[i].matches[j].result);
        }
    }
    wb.write(args.excel);
}

function populateTeams(teams, match) {//Put team in teams array if missing

    let t1idx = teams.findIndex(function (team) {

        if (team.name == match.t1) {
            return true;
        }
        else {
            return false;
        }
    });

    if (t1idx == -1) {
        let team = {
            name: match.t1,
            matches: []
        };
        teams.push(team);
    }

    let t2idx = teams.findIndex(function (team) {

        if (team.name == match.t2) {
            return true;
        }
        else {
            return false;
        }
    });

    if (t2idx == -1) {
        let team = {
            name: match.t2,
            matches: []
        };
        teams.push(team);
    }

}

function populateMatchesinTeams(teams, match)//put matches in appropriate team
{
    let t1idx = teams.findIndex(function (team) {

        if (team.name == match.t1) {
            return true;
        }
        else {
            return false;
        }
    });

    let team1 = teams[t1idx];
    team1.matches.push({
        vs: match.t2,
        selfScore: match.t1s,
        oppScore: match.t2s,
        result: match.result

    })

    let t2idx = teams.findIndex(function (team) {

        if (team.name == match.t2) {
            return true;
        }
        else {
            return false;
        }
    });

    let team2 = teams[t2idx];
    team2.matches.push({
        vs: match.t1,
        selfScore: match.t2s,
        oppScore: match.t1s,
        result: match.result

    })


}






