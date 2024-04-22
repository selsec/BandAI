function uniformOrder() {
    //start basics
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "Uniform Order";
    var sheet = ss.getSheetByName(sheetName);
    //check if sheet exists, if it does, clear it. If it doesn't create it
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
    } else {
        sheet.clear();
    }
    //declare count variables
    var smallCount = 0;
    var mediumCount = 0;
    var largeCount = 0;
    var xLargeCount = 0;
    var xxLargeCount = 0;
    var clarinetSmallCount = 0;
    var clarinetMediumCount = 0;
    var clarinetLargeCount = 0;
    var clarinetXLargeCount = 0;
    var clarinetXXLargeCount = 0;
    var saxophoneSmallCount = 0;
    var saxophoneMediumCount = 0;
    var saxophoneLargeCount = 0;
    var saxophoneXLargeCount = 0;
    var saxophoneXXLargeCount = 0;
    var trumpetSmallCount = 0;
    var trumpetMediumCount = 0;
    var trumpetLargeCount = 0;
    var trumpetXLargeCount = 0;
    var trumpetXXLargeCount = 0;
    var mellophoneSmallCount = 0;
    var mellophoneMediumCount = 0;
    var mellophoneLargeCount = 0;
    var mellophoneXLargeCount = 0;
    var mellophoneXXLargeCount = 0;
    var lowBrassSmallCount = 0;
    var lowBrassMediumCount = 0;
    var lowBrassLargeCount = 0;
    var lowBrassXLargeCount = 0;
    var lowBrassXXLargeCount = 0;
    var tubaSmallCount = 0;
    var tubaMediumCount = 0;
    var tubaLargeCount = 0;
    var tubaXLargeCount = 0;
    var tubaXXLargeCount = 0;
    var percussionSmallCount = 0;
    var percussionMediumCount = 0;
    var percussionLargeCount = 0;
    var percussionXLargeCount = 0;
    var percussionXXLargeCount = 0;
    var colorguardSmallCount = 0;
    var colorguardMediumCount = 0;
    var colorguardLargeCount = 0;
    var colorguardXLargeCount = 0;
    var colorguardXXLargeCount = 0;
    var fluteSmallCount = 0;
    var fluteMediumCount = 0;
    var fluteLargeCount = 0;
    var fluteXLargeCount = 0;
    var fluteXXLargeCount = 0;
    var marchingShoes55Count = 0;
    var marchingShoes6Count = 0;
    var marchingShoes65Count = 0;
    var marchingShoes7Count = 0;
    var marchingShoes75Count = 0;
    var marchingShoes8Count = 0;
    var marchingShoes85Count = 0;
    var marchingShoes9Count = 0;
    var marchingShoes95Count = 0;
    var marchingShoes10Count = 0;
    var marchingShoes105Count = 0;
    var marchingShoes11Count = 0;
    var marchingShoes115Count = 0;
    var marchingShoes12Count = 0;
    var marchingShoes125Count = 0;
    var marchingShoes13Count = 0;
    var marchingShoes135Count = 0;
    var marchingShoes14Count = 0;
    var marchingShoesOtherCount = 0;

    //format the sheet
    sheet.getRange("A1:M1").merge().setValue("Uniform Order").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("A1:M1").setBackground("#213483").setFontColor("#FFFFFF");
    sheet.getRange("A1:M1").setBorder(true, true, true, true, true, true);

    //create the show shirt order
    sheet.getRange("A2:B2").merge().setValue("Show Shirt Order").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("A2:B2").setBackground("#213483").setFontColor("#FFFFFF");
    sheet.getRange("A2:B2").setBorder(true, true, true, true, true, true);
    sheet.getRange("A3").setValue("S").setFontWeight("bold");
    sheet.getRange("A4").setValue("M").setFontWeight("bold");
    sheet.getRange("A5").setValue("L").setFontWeight("bold");  
    sheet.getRange("A6").setValue("XL").setFontWeight("bold");
    sheet.getRange("A7").setValue("XXL").setFontWeight("bold");
    sheet.getRange("A3:A7").setBorder(true, true, true, true, true, true);
    sheet.getRange("B3:B7").setBorder(true, true, true, true, true, true);

    //fill data for show shirt order
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Period Roster" && sheetName !== "Attendance" && sheetName !== "Uniform Order") {
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (a8Value === "s") {
                smallCount++;
            } else if (a8Value === "m") {
                mediumCount++;
            } else if (a8Value === "l") {
                largeCount++;
            } else if (a8Value === "xl") {
                xLargeCount++;
            } else if (a8Value === "xxl") {
                xxLargeCount++;
            }
            else {
                continue;
            }
        }
    }
    sheet.getRange("B3").setValue(smallCount);
    sheet.getRange("B4").setValue(mediumCount);
    sheet.getRange("B5").setValue(largeCount);
    sheet.getRange("B6").setValue(xLargeCount);
    sheet.getRange("B7").setValue(xxLargeCount);

    //create the Section Shirt Order
    sheet.getRange("D2:I2").merge().setValue("Section Shirt Order").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("D2:I2").setBackground("#213483").setFontColor("#FFFFFF");
    sheet.getRange("D2:I2").setBorder(true, true, true, true, true, true);
    sheet.getRange("D3").setValue("Flute").setFontWeight("bold");
    sheet.getRange("E3").setValue(fluteColor).setFontWeight("bold");
    sheet.getRange("D4").setValue("S").setFontWeight("bold");
    sheet.getRange("D5").setValue("M").setFontWeight("bold");
    sheet.getRange("D6").setValue("L").setFontWeight("bold");
    sheet.getRange("D7").setValue("XL").setFontWeight("bold");
    sheet.getRange("D8").setValue("XXL").setFontWeight("bold");
    sheet.getRange("D9").setValue("Clarinet").setFontWeight("bold");
    sheet.getRange("E9").setValue(clarinetColor).setFontWeight("bold");
    sheet.getRange("D10").setValue("S").setFontWeight("bold");
    sheet.getRange("D11").setValue("M").setFontWeight("bold");
    sheet.getRange("D12").setValue("L").setFontWeight("bold");
    sheet.getRange("D13").setValue("XL").setFontWeight("bold");
    sheet.getRange("D14").setValue("XXL").setFontWeight("bold");
    sheet.getRange("D15").setValue("Saxophone").setFontWeight("bold");
    sheet.getRange("E15").setValue(saxophoneColor).setFontWeight("bold");
    sheet.getRange("D16").setValue("S").setFontWeight("bold");
    sheet.getRange("D17").setValue("M").setFontWeight("bold");
    sheet.getRange("D18").setValue("L").setFontWeight("bold");
    sheet.getRange("D19").setValue("XL").setFontWeight("bold");
    sheet.getRange("D20").setValue("XXL").setFontWeight("bold");
    sheet.getRange("F3").setValue("Trumpet").setFontWeight("bold");
    sheet.getRange("G3").setValue(trumpetColor).setFontWeight("bold");
    sheet.getRange("F4").setValue("S").setFontWeight("bold");
    sheet.getRange("F5").setValue("M").setFontWeight("bold");
    sheet.getRange("F6").setValue("L").setFontWeight("bold");
    sheet.getRange("F7").setValue("XL").setFontWeight("bold");
    sheet.getRange("F8").setValue("XXL").setFontWeight("bold");
    sheet.getRange("F9").setValue("Mellophone").setFontWeight("bold");
    sheet.getRange("G9").setValue(mellophoneColor).setFontWeight("bold");
    sheet.getRange("F10").setValue("S").setFontWeight("bold");
    sheet.getRange("F11").setValue("M").setFontWeight("bold");
    sheet.getRange("F12").setValue("L").setFontWeight("bold");
    sheet.getRange("F13").setValue("XL").setFontWeight("bold");
    sheet.getRange("F14").setValue("XXL").setFontWeight("bold");
    sheet.getRange("F15").setValue("Low Brass").setFontWeight("bold");
    sheet.getRange("G15").setValue(lowBrassColor).setFontWeight("bold");
    sheet.getRange("F16").setValue("S").setFontWeight("bold");
    sheet.getRange("F17").setValue("M").setFontWeight("bold");
    sheet.getRange("F18").setValue("L").setFontWeight("bold");
    sheet.getRange("F19").setValue("XL").setFontWeight("bold");
    sheet.getRange("F20").setValue("XXL").setFontWeight("bold");
    sheet.getRange("H3").setValue("Tuba").setFontWeight("bold");
    sheet.getRange("I3").setValue(tubaColor).setFontWeight("bold");
    sheet.getRange("H4").setValue("S").setFontWeight("bold");
    sheet.getRange("H5").setValue("M").setFontWeight("bold");
    sheet.getRange("H6").setValue("L").setFontWeight("bold");
    sheet.getRange("H7").setValue("XL").setFontWeight("bold");
    sheet.getRange("H8").setValue("XXL").setFontWeight("bold");
    sheet.getRange("H9").setValue("Percussion").setFontWeight("bold");
    sheet.getRange("I9").setValue(percussionColor).setFontWeight("bold");
    sheet.getRange("H10").setValue("S").setFontWeight("bold");
    sheet.getRange("H11").setValue("M").setFontWeight("bold");
    sheet.getRange("H12").setValue("L").setFontWeight("bold");
    sheet.getRange("H13").setValue("XL").setFontWeight("bold");
    sheet.getRange("H14").setValue("XXL").setFontWeight("bold");
    sheet.getRange("H15").setValue("Colorguard").setFontWeight("bold");
    sheet.getRange("I15").setValue(colorguardColor).setFontWeight("bold");
    sheet.getRange("H16").setValue("S").setFontWeight("bold");
    sheet.getRange("H17").setValue("M").setFontWeight("bold");
    sheet.getRange("H18").setValue("L").setFontWeight("bold");
    sheet.getRange("H19").setValue("XL").setFontWeight("bold");
    sheet.getRange("H20").setValue("XXL").setFontWeight("bold");
    sheet.getRange("D3:I20").setBorder(true, true, true, true, true, true);

    //fill data for section shirt order
    
    //flute section order
    var sheets = ss.getSheets();
    
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E2").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (e2Value === "flute" && a8Value === "s") {
                fluteSmallCount++;
            } else if (e2Value === "flute" && a8Value === "m") {
                fluteMediumCount++;
            } else if (e2Value === "flute" && a8Value === "l") {
                fluteLargeCount++;
            } else if (e2Value === "flute" && a8Value === "xl") {
                fluteXLargeCount++;
            } else if (e2Value === "flute" && a8Value === "xxl") {
                fluteXXLargeCount++;
            }
        }
    }
    sheet.getRange("E4").setValue(fluteSmallCount);
    sheet.getRange("E5").setValue(fluteMediumCount);
    sheet.getRange("E6").setValue(fluteLargeCount);
    sheet.getRange("E7").setValue(fluteXLargeCount);
    sheet.getRange("E8").setValue(fluteXXLargeCount);
    
    //clarinet section order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E9").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (e2Value === "clarinet" && a8Value === "s") {
                clarinetSmallCount++;
            } else if (e2Value === "clarinet" && a8Value === "m") {
                clarinetMediumCount++;
            } else if (e2Value === "clarinet" && a8Value === "l") {
                clarinetLargeCount++;
            } else if (e2Value === "clarinet" && a8Value === "xl") {
                clarinetXLargeCount++;
            } else if (e2Value === "clarinet" && a8Value === "xxl") {
                clarinetXXLargeCount++;
            }
        }
    }
    sheet.getRange("E10").setValue(clarinetSmallCount);
    sheet.getRange("E11").setValue(clarinetMediumCount);
    sheet.getRange("E12").setValue(clarinetLargeCount);
    sheet.getRange("E13").setValue(clarinetXLargeCount);
    sheet.getRange("E14").setValue(clarinetXXLargeCount);
    //saxophone section order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E2").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (e2Value === "saxophone" && a8Value === "s") {
                saxophoneSmallCount++;
            } else if (e2Value === "saxophone" && a8Value === "m") {
                saxophoneMediumCount++;
            } else if (e2Value === "saxophone" && a8Value === "l") {
                saxophoneLargeCount++;
            } else if (e2Value === "saxophone" && a8Value === "xl") {
                saxophoneXLargeCount++;
            } else if (e2Value === "saxophone" && a8Value === "xxl") {
                saxophoneXXLargeCount++;
            }
        }
    }
    sheet.getRange("E16").setValue(saxophoneSmallCount);
    sheet.getRange("E17").setValue(saxophoneMediumCount);
    sheet.getRange("E18").setValue(saxophoneLargeCount);
    sheet.getRange("E19").setValue(saxophoneXLargeCount);
    sheet.getRange("E20").setValue(saxophoneXXLargeCount);
    //trumpet section order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E2").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (e2Value === "trumpet" && a8Value === "s") {
                trumpetSmallCount++;
            } else if (e2Value === "trumpet" && a8Value === "m") {
                trumpetMediumCount++;
            } else if (e2Value === "trumpet" && a8Value === "l") {
                trumpetLargeCount++;
            } else if (e2Value === "trumpet" && a8Value === "xl") {
                trumpetXLargeCount++;
            } else if (e2Value === "trumpet" && a8Value === "xxl") {
                trumpetXXLargeCount++;
            }
        }
    }
    sheet.getRange("G4").setValue(trumpetSmallCount);
    sheet.getRange("G5").setValue(trumpetMediumCount);
    sheet.getRange("G6").setValue(trumpetLargeCount);
    sheet.getRange("G7").setValue(trumpetXLargeCount);
    sheet.getRange("G8").setValue(trumpetXXLargeCount);
    //mellophone section order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E2").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (e2Value === "mellophone" && a8Value === "s") {
                mellophoneSmallCount++;
            } else if (e2Value === "mellophone" && a8Value === "m") {
                mellophoneMediumCount++;
            } else if (e2Value === "mellophone" && a8Value === "l") {
                mellophoneLargeCount++;
            } else if (e2Value === "mellophone" && a8Value === "xl") {
                mellophoneXLargeCount++;
            } else if (e2Value === "mellophone" && a8Value === "xxl") {
                mellophoneXXLargeCount++;
            }
        }
    }
    sheet.getRange("G10").setValue(mellophoneSmallCount);
    sheet.getRange("G11").setValue(mellophoneMediumCount);
    sheet.getRange("G12").setValue(mellophoneLargeCount);
    sheet.getRange("G13").setValue(mellophoneXLargeCount);
    sheet.getRange("G14").setValue(mellophoneXXLargeCount);
    //low brass section order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E2").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if ((e2Value === "trombone" || e2Value === "baritone") && a8Value === "s") {
                lowBrassSmallCount++;
            } else if ((e2Value === "trombone" || e2Value === "baritone") && a8Value === "m") {
                lowBrassMediumCount++;
            } else if ((e2Value === "trombone" || e2Value === "baritone") && a8Value === "l") {
                lowBrassLargeCount++;
            } else if ((e2Value === "trombone" || e2Value === "baritone") && a8Value === "xl") {
                lowBrassXLargeCount++;
            } else if ((e2Value === "trombone" || e2Value === "baritone") && a8Value === "xxl") {
                lowBrassXXLargeCount++;
            }
        } 
    }
    sheet.getRange("G16").setValue(lowBrassSmallCount);
    sheet.getRange("G17").setValue(lowBrassMediumCount);
    sheet.getRange("G18").setValue(lowBrassLargeCount);
    sheet.getRange("G19").setValue(lowBrassXLargeCount);
    sheet.getRange("G20").setValue(lowBrassXXLargeCount);
    //tuba section order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E2").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (e2Value === "tuba" && a8Value === "s") {
                tubaSmallCount++;
            } else if (e2Value === "tuba" && a8Value === "m") {
                tubaMediumCount++;
            } else if (e2Value === "tuba" && a8Value === "l") {
                tubaLargeCount++;
            } else if (e2Value === "tuba" && a8Value === "xl") {
                tubaXLargeCount++;
            } else if (e2Value === "tuba" && a8Value === "xxl") {
                tubaXXLargeCount++;
            }
        }
    }
    sheet.getRange("I4").setValue(tubaSmallCount);
    sheet.getRange("I5").setValue(tubaMediumCount);
    sheet.getRange("I6").setValue(tubaLargeCount);
    sheet.getRange("I7").setValue(tubaXLargeCount);
    sheet.getRange("I8").setValue(tubaXXLargeCount);
    //percussion section order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E2").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (e2Value === "percussion" && a8Value === "s") {
                percussionSmallCount++;
            } else if (e2Value === "percussion" && a8Value === "m") {
                percussionMediumCount++;
            } else if (e2Value === "percussion" && a8Value === "l") {
                percussionLargeCount++;
            } else if (e2Value === "percussion" && a8Value === "xl") {
                percussionXLargeCount++;
            } else if (e2Value === "percussion" && a8Value === "xxl") {
                percussionXXLargeCount++;
            }
        }
    }
    sheet.getRange("I10").setValue(percussionSmallCount);
    sheet.getRange("I11").setValue(percussionMediumCount);
    sheet.getRange("I12").setValue(percussionLargeCount);
    sheet.getRange("I13").setValue(percussionXLargeCount);
    sheet.getRange("I14").setValue(percussionXXLargeCount);
    //colorguard section order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E2").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (e2Value === "colorguard" && a8Value === "s") {
                colorguardSmallCount++;
            } else if (e2Value === "colorguard" && a8Value === "m") {
                colorguardMediumCount++;
            } else if (e2Value === "colorguard" && a8Value === "l") {
                colorguardLargeCount++;
            } else if (e2Value === "colorguard" && a8Value === "xl") {
                colorguardXLargeCount++;
            } else if (e2Value === "colorguard" && a8Value === "xxl") {
                colorguardXXLargeCount++;
            }
        }
    }
    sheet.getRange("I16").setValue(colorguardSmallCount);
    sheet.getRange("I17").setValue(colorguardMediumCount);
    sheet.getRange("I18").setValue(colorguardLargeCount);
    sheet.getRange("I19").setValue(colorguardXLargeCount);
    sheet.getRange("I20").setValue(colorguardXXLargeCount);
    
    //create the marching shoes format
    sheet.getRange("A9:B9").merge().setValue("Marching Shoes Order").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("A9:B9").setBackground("#213483").setFontColor("#FFFFFF");
    sheet.getRange("A9:B9").setBorder(true, true, true, true, true, true);
    sheet.getRange("A10").setValue("5.5").setFontWeight("bold");
    sheet.getRange("A11").setValue("6").setFontWeight("bold");
    sheet.getRange("A12").setValue("6.5").setFontWeight("bold");
    sheet.getRange("A13").setValue("7").setFontWeight("bold");
    sheet.getRange("A14").setValue("7.5").setFontWeight("bold");
    sheet.getRange("A15").setValue("8").setFontWeight("bold");
    sheet.getRange("A16").setValue("8.5").setFontWeight("bold");
    sheet.getRange("A17").setValue("9").setFontWeight("bold");
    sheet.getRange("A18").setValue("9.5").setFontWeight("bold");
    sheet.getRange("A19").setValue("10").setFontWeight("bold");
    sheet.getRange("A20").setValue("10.5").setFontWeight("bold");
    sheet.getRange("A21").setValue("11").setFontWeight("bold");
    sheet.getRange("A22").setValue("11.5").setFontWeight("bold");
    sheet.getRange("A23").setValue("12").setFontWeight("bold");
    sheet.getRange("A24").setValue("12.5").setFontWeight("bold");
    sheet.getRange("A25").setValue("13").setFontWeight("bold");
    sheet.getRange("A26").setValue("13.5").setFontWeight("bold");
    sheet.getRange("A27").setValue("14").setFontWeight("bold");
    sheet.getRange("A28:B28").merge().setValue("Other Sizes").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("A28:B28").setBackground("#213483").setFontColor("#FFFFFF");
    sheet.getRange("A28:B28").setBorder(true, true, true, true, true, true);
    
    //fill marching shoes data
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var a9Value = currentSheet.getRange("A9").getValue();
            if (a9Value === true) {
                var a8Value = currentSheet.getRange("A8").getValue().toString().toLowerCase;
                if (a8Value === "5.5") {
                    marchingShoes55Count = marchingShoes55Count + 1;
                } else if (a8Value === "6") {
                    marchingShoes6Count = marchingShoes6Count +1;
                }
                else if (a8Value === "6.5") {
                    marchingShoes65Count = marchingShoes65Count + 1;
                }
                else if (a8Value === "7") {
                    marchingShoes7Count = marchingShoes7Count + 1;
                }
                else if (a8Value === "7.5") {
                    marchingShoes75Count = marchingShoes75Count + 1;
                }
                else if (a8Value === "8") {
                    marchingShoes8Count = marchingShoes8Count + 1;
                }
                else if (a8Value === "8.5") {
                    marchingShoes85Count = marchingShoes85Count + 1;
                }
                else if (a8Value === "9") {
                    marchingShoes9Count = marchingShoes9Count + 1;
                }
                else if (a8Value === "9.5") {
                    marchingShoes95Count = marchingShoes95Count + 1;
                }
                else if (a8Value === "10") {
                    marchingShoes10Count = marchingShoes10Count + 1;
                }
                else if (a8Value === "10.5") {
                    marchingShoes105Count = marchingShoes105Count + 1;
                }
                else if (a8Value === "11") {
                    marchingShoes11Count = marchingShoes11Count + 1;
                }
                else if (a8Value === "11.5") {
                    marchingShoes115Count = marchingShoes115Count + 1;  
                }
                else if (a8Value === "12") {
                    marchingShoes12Count = function uniformOrder() {
    //start basics
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "Uniform Order";
    var sheet = ss.getSheetByName(sheetName);
    //check if sheet exists, if it does, clear it. If it doesn't create it
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
    } else {
        sheet.clear();
    }
    //declare count variables
    var smallCount = 0;
    var mediumCount = 0;
    var largeCount = 0;
    var xLargeCount = 0;
    var xxLargeCount = 0;
    var clarinetSmallCount = 0;
    var clarinetMediumCount = 0;
    var clarinetLargeCount = 0;
    var clarinetXLargeCount = 0;
    var clarinetXXLargeCount = 0;
    var saxophoneSmallCount = 0;
    var saxophoneMediumCount = 0;
    var saxophoneLargeCount = 0;
    var saxophoneXLargeCount = 0;
    var saxophoneXXLargeCount = 0;
    var trumpetSmallCount = 0;
    var trumpetMediumCount = 0;
    var trumpetLargeCount = 0;
    var trumpetXLargeCount = 0;
    var trumpetXXLargeCount = 0;
    var mellophoneSmallCount = 0;
    var mellophoneMediumCount = 0;
    var mellophoneLargeCount = 0;
    var mellophoneXLargeCount = 0;
    var mellophoneXXLargeCount = 0;
    var lowBrassSmallCount = 0;
    var lowBrassMediumCount = 0;
    var lowBrassLargeCount = 0;
    var lowBrassXLargeCount = 0;
    var lowBrassXXLargeCount = 0;
    var tubaSmallCount = 0;
    var tubaMediumCount = 0;
    var tubaLargeCount = 0;
    var tubaXLargeCount = 0;
    var tubaXXLargeCount = 0;
    var percussionSmallCount = 0;
    var percussionMediumCount = 0;
    var percussionLargeCount = 0;
    var percussionXLargeCount = 0;
    var percussionXXLargeCount = 0;
    var colorguardSmallCount = 0;
    var colorguardMediumCount = 0;
    var colorguardLargeCount = 0;
    var colorguardXLargeCount = 0;
    var colorguardXXLargeCount = 0;
    var fluteSmallCount = 0;
    var fluteMediumCount = 0;
    var fluteLargeCount = 0;
    var fluteXLargeCount = 0;
    var fluteXXLargeCount = 0;
    var marchingShoes55Count = 0;
    var marchingShoes6Count = 0;
    var marchingShoes65Count = 0;
    var marchingShoes7Count = 0;
    var marchingShoes75Count = 0;
    var marchingShoes8Count = 0;
    var marchingShoes85Count = 0;
    var marchingShoes9Count = 0;
    var marchingShoes95Count = 0;
    var marchingShoes10Count = 0;
    var marchingShoes105Count = 0;
    var marchingShoes11Count = 0;
    var marchingShoes115Count = 0;
    var marchingShoes12Count = 0;
    var marchingShoes125Count = 0;
    var marchingShoes13Count = 0;
    var marchingShoes135Count = 0;
    var marchingShoes14Count = 0;
    var marchingShoesOtherCount = 0;

    //format the sheet
    sheet.getRange("A1:M1").merge().setValue("Uniform Order").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("A1:M1").setBackground("#213483").setFontColor("#FFFFFF");
    sheet.getRange("A1:M1").setBorder(true, true, true, true, true, true);

    //create the show shirt order
    sheet.getRange("A2:B2").merge().setValue("Show Shirt Order").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("A2:B2").setBackground("#213483").setFontColor("#FFFFFF");
    sheet.getRange("A2:B2").setBorder(true, true, true, true, true, true);
    sheet.getRange("A3").setValue("S").setFontWeight("bold");
    sheet.getRange("A4").setValue("M").setFontWeight("bold");
    sheet.getRange("A5").setValue("L").setFontWeight("bold");  
    sheet.getRange("A6").setValue("XL").setFontWeight("bold");
    sheet.getRange("A7").setValue("XXL").setFontWeight("bold");
    sheet.getRange("A3:A7").setBorder(true, true, true, true, true, true);
    sheet.getRange("B3:B7").setBorder(true, true, true, true, true, true);

    //fill data for show shirt order
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Period Roster" && sheetName !== "Attendance" && sheetName !== "Uniform Order") {
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (a8Value === "s") {
                smallCount++;
            } else if (a8Value === "m") {
                mediumCount++;
            } else if (a8Value === "l") {
                largeCount++;
            } else if (a8Value === "xl") {
                xLargeCount++;
            } else if (a8Value === "xxl") {
                xxLargeCount++;
            }
            else {
                continue;
            }
        }
    }
    sheet.getRange("B3").setValue(smallCount);
    sheet.getRange("B4").setValue(mediumCount);
    sheet.getRange("B5").setValue(largeCount);
    sheet.getRange("B6").setValue(xLargeCount);
    sheet.getRange("B7").setValue(xxLargeCount);

    //create the Section Shirt Order
    sheet.getRange("D2:I2").merge().setValue("Section Shirt Order").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("D2:I2").setBackground("#213483").setFontColor("#FFFFFF");
    sheet.getRange("D2:I2").setBorder(true, true, true, true, true, true);
    sheet.getRange("D3").setValue("Flute").setFontWeight("bold");
    sheet.getRange("E3").setValue(fluteColor).setFontWeight("bold");
    sheet.getRange("D4").setValue("S").setFontWeight("bold");
    sheet.getRange("D5").setValue("M").setFontWeight("bold");
    sheet.getRange("D6").setValue("L").setFontWeight("bold");
    sheet.getRange("D7").setValue("XL").setFontWeight("bold");
    sheet.getRange("D8").setValue("XXL").setFontWeight("bold");
    sheet.getRange("D9").setValue("Clarinet").setFontWeight("bold");
    sheet.getRange("E9").setValue(clarinetColor).setFontWeight("bold");
    sheet.getRange("D10").setValue("S").setFontWeight("bold");
    sheet.getRange("D11").setValue("M").setFontWeight("bold");
    sheet.getRange("D12").setValue("L").setFontWeight("bold");
    sheet.getRange("D13").setValue("XL").setFontWeight("bold");
    sheet.getRange("D14").setValue("XXL").setFontWeight("bold");
    sheet.getRange("D15").setValue("Saxophone").setFontWeight("bold");
    sheet.getRange("E15").setValue(saxophoneColor).setFontWeight("bold");
    sheet.getRange("D16").setValue("S").setFontWeight("bold");
    sheet.getRange("D17").setValue("M").setFontWeight("bold");
    sheet.getRange("D18").setValue("L").setFontWeight("bold");
    sheet.getRange("D19").setValue("XL").setFontWeight("bold");
    sheet.getRange("D20").setValue("XXL").setFontWeight("bold");
    sheet.getRange("F3").setValue("Trumpet").setFontWeight("bold");
    sheet.getRange("G3").setValue(trumpetColor).setFontWeight("bold");
    sheet.getRange("F4").setValue("S").setFontWeight("bold");
    sheet.getRange("F5").setValue("M").setFontWeight("bold");
    sheet.getRange("F6").setValue("L").setFontWeight("bold");
    sheet.getRange("F7").setValue("XL").setFontWeight("bold");
    sheet.getRange("F8").setValue("XXL").setFontWeight("bold");
    sheet.getRange("F9").setValue("Mellophone").setFontWeight("bold");
    sheet.getRange("G9").setValue(mellophoneColor).setFontWeight("bold");
    sheet.getRange("F10").setValue("S").setFontWeight("bold");
    sheet.getRange("F11").setValue("M").setFontWeight("bold");
    sheet.getRange("F12").setValue("L").setFontWeight("bold");
    sheet.getRange("F13").setValue("XL").setFontWeight("bold");
    sheet.getRange("F14").setValue("XXL").setFontWeight("bold");
    sheet.getRange("F15").setValue("Low Brass").setFontWeight("bold");
    sheet.getRange("G15").setValue(lowBrassColor).setFontWeight("bold");
    sheet.getRange("F16").setValue("S").setFontWeight("bold");
    sheet.getRange("F17").setValue("M").setFontWeight("bold");
    sheet.getRange("F18").setValue("L").setFontWeight("bold");
    sheet.getRange("F19").setValue("XL").setFontWeight("bold");
    sheet.getRange("F20").setValue("XXL").setFontWeight("bold");
    sheet.getRange("H3").setValue("Tuba").setFontWeight("bold");
    sheet.getRange("I3").setValue(tubaColor).setFontWeight("bold");
    sheet.getRange("H4").setValue("S").setFontWeight("bold");
    sheet.getRange("H5").setValue("M").setFontWeight("bold");
    sheet.getRange("H6").setValue("L").setFontWeight("bold");
    sheet.getRange("H7").setValue("XL").setFontWeight("bold");
    sheet.getRange("H8").setValue("XXL").setFontWeight("bold");
    sheet.getRange("H9").setValue("Percussion").setFontWeight("bold");
    sheet.getRange("I9").setValue(percussionColor).setFontWeight("bold");
    sheet.getRange("H10").setValue("S").setFontWeight("bold");
    sheet.getRange("H11").setValue("M").setFontWeight("bold");
    sheet.getRange("H12").setValue("L").setFontWeight("bold");
    sheet.getRange("H13").setValue("XL").setFontWeight("bold");
    sheet.getRange("H14").setValue("XXL").setFontWeight("bold");
    sheet.getRange("H15").setValue("Colorguard").setFontWeight("bold");
    sheet.getRange("I15").setValue(colorguardColor).setFontWeight("bold");
    sheet.getRange("H16").setValue("S").setFontWeight("bold");
    sheet.getRange("H17").setValue("M").setFontWeight("bold");
    sheet.getRange("H18").setValue("L").setFontWeight("bold");
    sheet.getRange("H19").setValue("XL").setFontWeight("bold");
    sheet.getRange("H20").setValue("XXL").setFontWeight("bold");
    sheet.getRange("D3:I20").setBorder(true, true, true, true, true, true);

    //fill data for section shirt order
    
    //flute section order
    var sheets = ss.getSheets();
    
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E2").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (e2Value === "flute" && a8Value === "s") {
                fluteSmallCount++;
            } else if (e2Value === "flute" && a8Value === "m") {
                fluteMediumCount++;
            } else if (e2Value === "flute" && a8Value === "l") {
                fluteLargeCount++;
            } else if (e2Value === "flute" && a8Value === "xl") {
                fluteXLargeCount++;
            } else if (e2Value === "flute" && a8Value === "xxl") {
                fluteXXLargeCount++;
            }
        }
    }
    sheet.getRange("E4").setValue(fluteSmallCount);
    sheet.getRange("E5").setValue(fluteMediumCount);
    sheet.getRange("E6").setValue(fluteLargeCount);
    sheet.getRange("E7").setValue(fluteXLargeCount);
    sheet.getRange("E8").setValue(fluteXXLargeCount);
    
    //clarinet section order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E9").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (e2Value === "clarinet" && a8Value === "s") {
                clarinetSmallCount++;
            } else if (e2Value === "clarinet" && a8Value === "m") {
                clarinetMediumCount++;
            } else if (e2Value === "clarinet" && a8Value === "l") {
                clarinetLargeCount++;
            } else if (e2Value === "clarinet" && a8Value === "xl") {
                clarinetXLargeCount++;
            } else if (e2Value === "clarinet" && a8Value === "xxl") {
                clarinetXXLargeCount++;
            }
        }
    }
    sheet.getRange("E10").setValue(clarinetSmallCount);
    sheet.getRange("E11").setValue(clarinetMediumCount);
    sheet.getRange("E12").setValue(clarinetLargeCount);
    sheet.getRange("E13").setValue(clarinetXLargeCount);
    sheet.getRange("E14").setValue(clarinetXXLargeCount);
    //saxophone section order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E2").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (e2Value === "saxophone" && a8Value === "s") {
                saxophoneSmallCount++;
            } else if (e2Value === "saxophone" && a8Value === "m") {
                saxophoneMediumCount++;
            } else if (e2Value === "saxophone" && a8Value === "l") {
                saxophoneLargeCount++;
            } else if (e2Value === "saxophone" && a8Value === "xl") {
                saxophoneXLargeCount++;
            } else if (e2Value === "saxophone" && a8Value === "xxl") {
                saxophoneXXLargeCount++;
            }
        }
    }
    sheet.getRange("E16").setValue(saxophoneSmallCount);
    sheet.getRange("E17").setValue(saxophoneMediumCount);
    sheet.getRange("E18").setValue(saxophoneLargeCount);
    sheet.getRange("E19").setValue(saxophoneXLargeCount);
    sheet.getRange("E20").setValue(saxophoneXXLargeCount);
    //trumpet section order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E2").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (e2Value === "trumpet" && a8Value === "s") {
                trumpetSmallCount++;
            } else if (e2Value === "trumpet" && a8Value === "m") {
                trumpetMediumCount++;
            } else if (e2Value === "trumpet" && a8Value === "l") {
                trumpetLargeCount++;
            } else if (e2Value === "trumpet" && a8Value === "xl") {
                trumpetXLargeCount++;
            } else if (e2Value === "trumpet" && a8Value === "xxl") {
                trumpetXXLargeCount++;
            }
        }
    }
    sheet.getRange("G4").setValue(trumpetSmallCount);
    sheet.getRange("G5").setValue(trumpetMediumCount);
    sheet.getRange("G6").setValue(trumpetLargeCount);
    sheet.getRange("G7").setValue(trumpetXLargeCount);
    sheet.getRange("G8").setValue(trumpetXXLargeCount);
    //mellophone section order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E2").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (e2Value === "mellophone" && a8Value === "s") {
                mellophoneSmallCount++;
            } else if (e2Value === "mellophone" && a8Value === "m") {
                mellophoneMediumCount++;
            } else if (e2Value === "mellophone" && a8Value === "l") {
                mellophoneLargeCount++;
            } else if (e2Value === "mellophone" && a8Value === "xl") {
                mellophoneXLargeCount++;
            } else if (e2Value === "mellophone" && a8Value === "xxl") {
                mellophoneXXLargeCount++;
            }
        }
    }
    sheet.getRange("G10").setValue(mellophoneSmallCount);
    sheet.getRange("G11").setValue(mellophoneMediumCount);
    sheet.getRange("G12").setValue(mellophoneLargeCount);
    sheet.getRange("G13").setValue(mellophoneXLargeCount);
    sheet.getRange("G14").setValue(mellophoneXXLargeCount);
    //low brass section order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E2").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if ((e2Value === "trombone" || e2Value === "baritone") && a8Value === "s") {
                lowBrassSmallCount++;
            } else if ((e2Value === "trombone" || e2Value === "baritone") && a8Value === "m") {
                lowBrassMediumCount++;
            } else if ((e2Value === "trombone" || e2Value === "baritone") && a8Value === "l") {
                lowBrassLargeCount++;
            } else if ((e2Value === "trombone" || e2Value === "baritone") && a8Value === "xl") {
                lowBrassXLargeCount++;
            } else if ((e2Value === "trombone" || e2Value === "baritone") && a8Value === "xxl") {
                lowBrassXXLargeCount++;
            }
        } 
    }
    sheet.getRange("G16").setValue(lowBrassSmallCount);
    sheet.getRange("G17").setValue(lowBrassMediumCount);
    sheet.getRange("G18").setValue(lowBrassLargeCount);
    sheet.getRange("G19").setValue(lowBrassXLargeCount);
    sheet.getRange("G20").setValue(lowBrassXXLargeCount);
    //tuba section order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E2").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (e2Value === "tuba" && a8Value === "s") {
                tubaSmallCount++;
            } else if (e2Value === "tuba" && a8Value === "m") {
                tubaMediumCount++;
            } else if (e2Value === "tuba" && a8Value === "l") {
                tubaLargeCount++;
            } else if (e2Value === "tuba" && a8Value === "xl") {
                tubaXLargeCount++;
            } else if (e2Value === "tuba" && a8Value === "xxl") {
                tubaXXLargeCount++;
            }
        }
    }
    sheet.getRange("I4").setValue(tubaSmallCount);
    sheet.getRange("I5").setValue(tubaMediumCount);
    sheet.getRange("I6").setValue(tubaLargeCount);
    sheet.getRange("I7").setValue(tubaXLargeCount);
    sheet.getRange("I8").setValue(tubaXXLargeCount);
    //percussion section order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E2").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (e2Value === "percussion" && a8Value === "s") {
                percussionSmallCount++;
            } else if (e2Value === "percussion" && a8Value === "m") {
                percussionMediumCount++;
            } else if (e2Value === "percussion" && a8Value === "l") {
                percussionLargeCount++;
            } else if (e2Value === "percussion" && a8Value === "xl") {
                percussionXLargeCount++;
            } else if (e2Value === "percussion" && a8Value === "xxl") {
                percussionXXLargeCount++;
            }
        }
    }
    sheet.getRange("I10").setValue(percussionSmallCount);
    sheet.getRange("I11").setValue(percussionMediumCount);
    sheet.getRange("I12").setValue(percussionLargeCount);
    sheet.getRange("I13").setValue(percussionXLargeCount);
    sheet.getRange("I14").setValue(percussionXXLargeCount);
    //colorguard section order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e2Value = currentSheet.getRange("E2").getValue().toString().toLowerCase();
            var a8Value = currentSheet.getRange("C8").getValue().toString().toLowerCase();
            if (e2Value === "colorguard" && a8Value === "s") {
                colorguardSmallCount++;
            } else if (e2Value === "colorguard" && a8Value === "m") {
                colorguardMediumCount++;
            } else if (e2Value === "colorguard" && a8Value === "l") {
                colorguardLargeCount++;
            } else if (e2Value === "colorguard" && a8Value === "xl") {
                colorguardXLargeCount++;
            } else if (e2Value === "colorguard" && a8Value === "xxl") {
                colorguardXXLargeCount++;
            }
        }
    }
    sheet.getRange("I16").setValue(colorguardSmallCount);
    sheet.getRange("I17").setValue(colorguardMediumCount);
    sheet.getRange("I18").setValue(colorguardLargeCount);
    sheet.getRange("I19").setValue(colorguardXLargeCount);
    sheet.getRange("I20").setValue(colorguardXXLargeCount);
    
    //create the marching shoes format
    sheet.getRange("A9:B9").merge().setValue("Marching Shoes Order").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("A9:B9").setBackground("#213483").setFontColor("#FFFFFF");
    sheet.getRange("A9:B9").setBorder(true, true, true, true, true, true);
    sheet.getRange("A10").setValue("5.5").setFontWeight("bold");
    sheet.getRange("A11").setValue("6").setFontWeight("bold");
    sheet.getRange("A12").setValue("6.5").setFontWeight("bold");
    sheet.getRange("A13").setValue("7").setFontWeight("bold");
    sheet.getRange("A14").setValue("7.5").setFontWeight("bold");
    sheet.getRange("A15").setValue("8").setFontWeight("bold");
    sheet.getRange("A16").setValue("8.5").setFontWeight("bold");
    sheet.getRange("A17").setValue("9").setFontWeight("bold");
    sheet.getRange("A18").setValue("9.5").setFontWeight("bold");
    sheet.getRange("A19").setValue("10").setFontWeight("bold");
    sheet.getRange("A20").setValue("10.5").setFontWeight("bold");
    sheet.getRange("A21").setValue("11").setFontWeight("bold");
    sheet.getRange("A22").setValue("11.5").setFontWeight("bold");
    sheet.getRange("A23").setValue("12").setFontWeight("bold");
    sheet.getRange("A24").setValue("12.5").setFontWeight("bold");
    sheet.getRange("A25").setValue("13").setFontWeight("bold");
    sheet.getRange("A26").setValue("13.5").setFontWeight("bold");
    sheet.getRange("A27").setValue("14").setFontWeight("bold");
    sheet.getRange("A28:B28").merge().setValue("Other Sizes").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("A28:B28").setBackground("#213483").setFontColor("#FFFFFF");
    sheet.getRange("A28:B28").setBorder(true, true, true, true, true, true);
    
    //fill marching shoes data
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var a9Value = currentSheet.getRange("A9").getValue();
            if (a9Value === true) {
                var a8Value = currentSheet.getRange("A8").getValue().toString().toLowerCase;
                if (a8Value === "5.5") {
                    marchingShoes55Count = marchingShoes55Count + 1;
                } else if (a8Value === "6") {
                    marchingShoes6Count = marchingShoes6Count +1;
                }
                else if (a8Value === "6.5") {
                    marchingShoes65Count = marchingShoes65Count + 1;
                }
                else if (a8Value === "7") {
                    marchingShoes7Count = marchingShoes7Count + 1;
                }
                else if (a8Value === "7.5") {
                    marchingShoes75Count = marchingShoes75Count + 1;
                }
                else if (a8Value === "8") {
                    marchingShoes8Count = marchingShoes8Count + 1;
                }
                else if (a8Value === "8.5") {
                    marchingShoes85Count = marchingShoes85Count + 1;
                }
                else if (a8Value === "9") {
                    marchingShoes9Count = marchingShoes9Count + 1;
                }
                else if (a8Value === "9.5") {
                    marchingShoes95Count = marchingShoes95Count + 1;
                }
                else if (a8Value === "10") {
                    marchingShoes10Count = marchingShoes10Count + 1;
                }
                else if (a8Value === "10.5") {
                    marchingShoes105Count = marchingShoes105Count + 1;
                }
                else if (a8Value === "11") {
                    marchingShoes11Count = marchingShoes11Count + 1;
                }
                else if (a8Value === "11.5") {
                    marchingShoes115Count = marchingShoes115Count + 1;
                }
                else if (a8Value === "12") {
                    marchingShoes12Count = marchingShoes12Count + 1;
                }
                else if (a8Value === "12.5") {
                    marchingShoes125Count  = marchingShoes125Count + 1;
                }
                else if (a8Value === "13") {
                    marchingShoes13Count = marchingShoes13Count + 1;
                }
                else if (a8Value === "13.5") {
                    marchingShoes135Count = marchingShoes135Count + 1;
                }
                else if (a8Value === "14") {
                    marchingShoes14Count = marchingShoes14Count + 1;
                }
                else {
                    marchingShoesOtherCount = marchingShoesOtherCount + ", " + a8Value;
                }
            }
            sheet.getRange("B10").setValue(marchingShoes55Count);
            sheet.getRange("B11").setValue(marchingShoes6Count);
            sheet.getRange("B12").setValue(marchingShoes65Count);
            sheet.getRange("B13").setValue(marchingShoes7Count);
            sheet.getRange("B14").setValue(marchingShoes75Count);
            sheet.getRange("B15").setValue(marchingShoes8Count);
            sheet.getRange("B16").setValue(marchingShoes85Count);
            sheet.getRange("B17").setValue(marchingShoes9Count);
            sheet.getRange("B18").setValue(marchingShoes95Count);
            sheet.getRange("B19").setValue(marchingShoes10Count);
            sheet.getRange("B20").setValue(marchingShoes105Count);
            sheet.getRange("B21").setValue(marchingShoes11Count);
            sheet.getRange("B22").setValue(marchingShoes115Count);
            sheet.getRange("B23").setValue(marchingShoes12Count);
            sheet.getRange("B24").setValue(marchingShoes125Count);
            sheet.getRange("B25").setValue(marchingShoes13Count);
            sheet.getRange("B26").setValue(marchingShoes135Count);
            sheet.getRange("B27").setValue(marchingShoes14Count);
            sheet.getRange("B28").setValue(marchingShoesOtherCount);
            

        }
        //bibbers order format
    }
                    }
                }
            }
        }
    }
}
