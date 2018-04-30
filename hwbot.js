//Requirements
var nodemailer = require('nodemailer');

//If you want to send homework from friday during the weekends
override = false;

//List of classes
teacherlist = ["GenericMathTeacher"]

//What you want the classes to be called (1st in the teacherlist will be the 1st here)
subjectlist = ["Math"]


//Style sheet (css) if you want to format the email.
stylesheet = "<style>.pointlist {background-color: white; border-radius: 5px; padding-left: 5px; padding-top: 5px; padding-bottom: 5px; padding-right: 5px;}  .wrapper {background-color: #fff8dd; border-radius: 5px; margin-top: 5px; margin-bottom: 5px; margin-left: 5px; margin-right: 5px; padding-left: 5px; padding-top: 5px; padding-bottom: 5px; padding-right: 5px;}   .classheader {background-color: #0061ff;color: white;border-radius: 5px;margin-top: 5px;margin-bottom: 5px;text-align: center; font-weight: 4px}</style>"

//People you want to send it to.
mailing_list = ["myRandomEmail@site.com"];


function SendEmail(text, subject) {
    var transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: 'MyEmail@site.com', // Your email
            pass: 'MyPass' // Your password
        }
    });

    var mailOptions = {
        from: 'Homework Bot',
        to: mailing_list.join(", "),
        subject: subject,
        html: text
    };

    transporter.sendMail(mailOptions, function(error, info) {
        if (error) {
            console.log(error);
        } else {
            console.log('Email sent: ' + info.response);
        }
    });

}


function Sheet(Column, teacherlist, inv) {

    var d = new Date();
    var weekday = new Array(7);
    weekday[0] = "January";
    weekday[1] = "February";
    weekday[2] = "March";
    weekday[3] = "April";
    weekday[4] = "May";
    weekday[5] = "June";
    weekday[6] = "July";
    weekday[7] = "August";
    weekday[8] = "September";
    weekday[9] = "October";
    weekday[10] = "November";
    weekday[11] = "December";


    var n = weekday[d.getMonth()];
    EmailSubject = "Homework for " + n + " " + d.getDate() + ", " + d.getFullYear();
    invmessage = "";
    if (inv == true) {
        invmessage = "<hr><i>Homework manually sent</i><hr>"
    }
    FullMSG = stylesheet + "<b>Homework for " + n + " " + d.getDate() + ", " + d.getFullYear() + ".</b> " + invmessage + " <br> <br><div class='wrapper'>"



    var GoogleSpreadsheet = require('google-spreadsheet');
    var async = require('async');

    // spreadsheet key is the long id in the sheets URL - barrington 7th
    var doc = new GoogleSpreadsheet('1q5Fl8vyDQ2Us-vpKX-J3SS8dXGwkK2axiLSEGmv5dIc');
    var sheet;

    async.series([
        function setAuth(step) {
            var creds = require('./DiscordSheet.json');
            var creds_json = {
                client_email: 'sheet-310@discordsheet-189222.iam.gserviceaccount.com',
                private_key: 'MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCv5doRPte/wv29hTKamcA+qa1ETZRI5xSXIGraVFU6k8vhPWZGPkVujYFZTB1hoQzdG8YCr5TASHyxNwBi7/+ig1p5m13uI9YEHP84dikVvCVqfnrxc3gLhotwjVARnCvQwXEFVws7e+DFHX763LqqXuUfMgZ3InjxVDLhJeCRWFf+T3urDYSSJBp5w3YFzKAVipRP+A8tFqeeCQEUR2p4f2Z8DhFZabW/hLdT1ZejOWfR05N2f6ekIOP3ljsyEFZHy5VnKbL4k8DfUy1wqJHbhTgs++7UA9Tvx/4QBTcQJ0pdTGnaABCgK8i1GYjO8ziDwbxI3hGyDt8sAFrujfh1AgMBAAECggEAJkz0Hg+zPlXo1uDfOzMx3WMob4g9+t7YSKz+H1LAKSxf90hdkfuCtlcZHHbaqfS9vpKK2/BcAT93aUYa7zAnl30nEPYv7of+lLA0sZhnC0zHz+fBAPy93bKa/7PBhMge8UqBv+Ird7UaZQluahTwj2Lw3IlRx5SpxZCqMITFIJiBrY2tAjoHIbrOaGTNe4NfUG87MkUepSIr2wthNVHIq4m3rDGIrQQr8vtl2491EHl0YWKz9FR/JBTNu85JemWDBjATExTOCMw5HXIiloyu91eoeZVznpoOmhoKIGTtUPOK4YXaN7+fnB1HRSy4WJZU68veEUnYcLPL909wZcgNOQKBgQDfZ04VKVYP3nugQ++bWpEazB11ocqj+ocOHZoS3q6BVyw8F+gZDzMjW2+Ep3Xlwletq/NxOcFTPfZ3Y/fD84RCm8BZKnsTxP1buAnFrT/85rQFKuI+He1QjTqeZl1LvJOomgyrfWllSHf/cgKByF5fwZFk8VwnN1n4J7WKHjOrrQKBgQDJkBaEu11CTiq99HPZXx9lDZs/q8Ns2jQgcCmbqXtK+ZdQTH3SUfmGTS+V0PgUOTuo+UyO14jdwoOtEkVcLxQ/V4wl0sv+4AVjTFYMKnKa8EtSFXsQ+ADQ1Lmilj++6SBv2X1cduZJeEX3tUP7tsCXgrWiLYRw+GUa14IYxPiY6QKBgQCF8wv3Vjya8TxQ4MsG7Cu61I4JXQQChBF8XjVmgQxC0wDb2z234Mw5s/ZOpJXlODyYDlS+G/IVtj5UYaVKYXV49qhDDlyTgvaiituZIgMO4UkpHAhhVFJZjQSLuzbPVXd3jT5xiJWsO+JvUG2+YWRsp8REsQ8fGGoih7Sq5ub2VQKBgGGUBsLFLfW+f8SsBjWSblFuj9z4G0ikLh4SDqKUHuMCB7XRAgyCsOjKjyHZI3Au8OnxnpV8VH7+/t4XfUqOZB/yEx/wd99wtksHFpUXK5pEgEJBse1aEpMGmMPUNxIGLDTZtm3ABeZFeqHbuAiwxXXTynizzm0eY1vmPs4c9TiZAoGBAL4Q9l0yXjG13VvSKUXwaTK+KLCdcyTKXdrlx+jYPAjD2j/LojfGzwudTe+6v487tZHy7x03YY6D+nqAePpBSwDoee6z1w57XzcB3Bn6hd50rTIyaYnmjhworVeY+szsIkY89cVew6Y851TyX7HV9fodOoVJiLP9LmpnBS1ifOuv'
            }

            doc.useServiceAccountAuth(creds, step);
        },
        function getInfoAndWorksheets(step) {
            doc.getInfo(function(err, info) {

                if (err) {
                    console.log('Error: ' + err);
                }


                NewGradesSheet = info.worksheets[0];

                step();



            });
        },


        function workingWithCells(step) {


            console.log("--> HWbot active");
            console.log("Fetching data...");


            NewGradesSheet.getCells({
                'min-row': 1,
                'max-row': 80,
                'return-empty': true
            }, function(err, cells) {




                i = 0;
                while (i < cells.length) {
                    f = 0;
                    while (f < teacherlist.length) {

                        if (cells[i].value == teacherlist[f]) {
                            cellData = cells[i + Column].value

                            if (cells[i + Column].value == "NH" || cells[i + Column].value == "No Homework") {
                                cellData = "No homework."
                            }
                            if (cells[i + Column].value == null || cells[i + Column].value == "" || cells[i + Column].value == " ") {
                                cellData = "No info found."
                            }

                            hw = "<div class='pointlist'>";
                            n = 0;

                            if (cellData != undefined) {
                                points = cellData.split("\n");
                                while (n < points.length) {
                                    hw = hw + "<b>-</b> " + points[n] + "<br>"
                                    n++;
                                }
                                hw = hw + "</div>"
                            } else {
                                hw = "No homework"
                            }

                            if (cellData != "No homework.") {
                                //format
                                FullMSG = FullMSG + "<div class='classheader'>" + subjectlist[f] + "</div>" + hw + "<br> ";
                            }




                        }
                        f++;
                    }




                    i++;
                }
                if (inv == false) {
                    //This is the stuff you want to add at the end.
                    FullMSG = FullMSG + "<div class='classheader'>Russian (Math)</div> <br> <br><div class='classheader'>Chemistry</div> <br> <br> <div class='classheader'>Russian</div> <br> <br> <div class='classheader'>Guitar</div> <br> <br> </div><br> <br> <hr> <i> This message is sent automatically </i>"
                }
                console.log("Data fetched --> Sending (This may take about 5 seconds)\n \n \n \n \n")
                setTimeout(() => {
                    SendEmail(FullMSG, EmailSubject);

                }, 5000)




            });




            step();



        }

    ], function(err) {
        if (err) {
            console.log('Error: ' + err);
        }
    });




}




date = new Date();
dt = date.getDay();
if (date.getDay() == 0 || date.getDay() == 6) {
    if (override == true) {
        console.log("Invalid day of the week --> No homework on weekends")
        console.log("Override --> Fetching homework for Friday")
        dt = 5;
        Sheet(dt, teacherlist, true);
    } else {
        console.log("Invalid day of the week --> No homework on weekends")
    }
} else {
    Sheet(date.getDay(), teacherlist, false);
}