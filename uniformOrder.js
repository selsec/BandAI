// Description: This script will create a new sheet called "Uniform Order" and populate it with the uniform order data from the individual sheets.

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
    sheet.getRange("A9:B28").setBorder(true, true, true, true, true, true);
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
            var b9value = currentSheet.getRange("A9").getValue();
            if (b9value === true) {
                var a8Value = currentSheet.getRange("A8").getValue().toString().toLowerCase();
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
                    marchingShoes125Count = marchingShoes125Count + 1;
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
                    if (marchingShoesOtherCount === "") {
                        marchingShoesOtherCount = a8Value;
                    } else {
                    marchingShoesOtherCount += ", " + a8Value;
                }
                }
            }
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
    sheet.getRange("A29").setValue(marchingShoesOtherCount);

    //bibbers format
    sheet.getRange("k2:l2").merge().setValue("Bibbers Order").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("k2:l2").setBackground("#213483").setFontColor("#FFFFFF");
    sheet.getRange("k2:l2").setBorder(true, true, true, true, true, true);
    sheet.getRange("K3:L11").setBorder(true, true, true, true, true, true).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange("K3").setValue("2XS");
    sheet.getRange("K4").setValue("XS");
    sheet.getRange("K5").setValue("S");
    sheet.getRange("K6").setValue("M");
    sheet.getRange("K7").setValue("L");
    sheet.getRange("K8").setValue("XL");
    sheet.getRange("K9").setValue("2XL");
    sheet.getRange("K10").setValue("3XL");
    sheet.getRange("K11").setValue("4XL");

    //bibbers order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var b9value = currentSheet.getRange("B9").getValue();
            if (b9value === true) {
                var i8hips = currentSheet.getRange("i8").getValue().toString().toLowerCase();
                if (i8hips === "27" || i8hips === "28" || i8hips === '29'){
                    bibbers2xsCount = bibbers2xsCount + 1;
                } else if (i8hips === "30" || i8hips === "31" || i8hips === "32" || i8hips === "33"){
                    bibbersXSCount = bibbersXSCount + 1;
                } else if (i8hips === "34" || i8hips === "35" || i8hips === "36" || i8hips === "37"){
                    bibbersSCount = bibbersSCount + 1;
                } else if (i8hips === "39" || i8hips === "40" || i8hips === "41" || i8hips === "42"){ 
                    bibbersMCount = bibbersMCount + 1;
                } else if (i8hips === "43" || i8hips === "44"){ 
                    bibbersLCount = bibbersLCount + 1;
                } else if (i8hips === "45" || i8hips === "46" || i8hips === "47" || i8hips === "48"){ 
                    bibbersXLCount = bibbersXLCount + 1;
                } else if (i8hips === "49" || i8hips === "50" || i8hips === "51" || i8hips === "52"){ 
                    bibbersXXLCount = bibbersXXLCount + 1;
                } else if (i8hips === "53" || i8hips === "54" || i8hips === "55" || i8hips === "56"){ 
                    bibbersXXXLCount = bibbersXXXLCount + 1;
                } else if (i8hips === "57" || i8hips === "58" || i8hips === "59" || i8hips === "60"){ 
                    bibbersXXXXLCount = bibbersXXXXLCount + 1;
                }
            }
        }
        
    }
    sheet.getRange("L3").setValue(bibbers2xsCount);
    sheet.getRange("L4").setValue(bibbersXSCount);
    sheet.getRange("L5").setValue(bibbersSCount);
    sheet.getRange("L6").setValue(bibbersMCount);
    sheet.getRange("L7").setValue(bibbersLCount);
    sheet.getRange("L8").setValue(bibbersXLCount);
    sheet.getRange("L9").setValue(bibbersXXLCount);
    sheet.getRange("L10").setValue(bibbersXXXLCount);
    sheet.getRange("L11").setValue(bibbersXXXXLCount);

    //gloves format
    sheet.getRange("d22:e22").merge().setValue("Gloves Order").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("d22:e22").setBackground("#213483").setFontColor("#FFFFFF");
    sheet.getRange("d22:e22").setBorder(true, true, true, true, true, true);
    sheet.getRange("d23:e27").setBorder(true, true, true, true, true, true).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange("d23").setValue("XS");
    sheet.getRange("d24").setValue("S");
    sheet.getRange("d25").setValue("M");
    sheet.getRange("d26").setValue("L");
    sheet.getRange("d27").setValue("XL");

    //gloves order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var f9value = currentSheet.getRange("H10").getValue();
            if (f9value === true) {
                var i10value = currentSheet.getRange("D8").getValue().toString().toLowerCase();
                if (i10value === "xs") {
                    glovesXSCount = glovesXSCount + 1;
                } else if (i10value === "s") {
                    glovesSCount = glovesSCount + 1;
                } else if (i10value === "m") {
                    glovesMCount = glovesMCount + 1;
                } else if (i10value === "l") {
                    glovesLCount = glovesLCount + 1;
                } else if (i10value === "xl") {
                    glovesXLCount = glovesXLCount + 1;
                }
            }
        }
    }
    sheet.getRange("E23").setValue(glovesXSCount);
    sheet.getRange("E24").setValue(glovesSCount);
    sheet.getRange("E25").setValue(glovesMCount);
    sheet.getRange("E26").setValue(glovesLCount);
    sheet.getRange("E27").setValue(glovesXLCount);

    //concert dress format
    sheet.getRange("f22:g22").merge().setValue("Concert Dress Order").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("f22:g22").setBackground("#213483").setFontColor("#FFFFFF");
    sheet.getRange("f22:g22").setBorder(true, true, true, true, true, true);
    sheet.getRange("f23:g37").setBorder(true, true, true, true, true, true).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange("f23").setValue("Size 0");
    sheet.getRange("f24").setValue("Size 2");
    sheet.getRange("f25").setValue("Size 4");
    sheet.getRange("f26").setValue("Size 6");
    sheet.getRange("f27").setValue("Size 8");
    sheet.getRange("f28").setValue("Size 10");
    sheet.getRange("f29").setValue("Size 12");
    sheet.getRange("f30").setValue("Size 14");
    sheet.getRange("f31").setValue("Size 16");
    sheet.getRange("f32").setValue("Size 18");
    sheet.getRange("f33").setValue("Size 20");
    sheet.getRange("f34").setValue("Size 22");
    sheet.getRange("f35").setValue("Size 24");
    sheet.getRange("f36").setValue("Size 26");
    sheet.getRange("f37").setValue("Size 28");

    //concert dress order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var f9value = currentSheet.getRange("f9").getValue();
            if (f9value === true) {
                var chest = currentSheet.getRange("g8").getValue().toString().toLowerCase();
                var waist = currentSheet.getRange("h8").getValue().toString().toLowerCase();
                var hips = currentSheet.getRange("i8").getValue().toString().toLowerCase();
                if (chest >= 29.5 && chest <= 31 && waist >= 20.5 && waist <= 22 && hips >= 30.5 && hips <= 32) {
                    size0Count++;
                  } else if (chest >= 30.5 && chest <= 33 && waist >= 22.5 && waist <= 24 && hips >= 32.5 && hips <= 34) {
                    size2Count++;
                  }
                  else if (chest >= 32.5 && chest <= 35 && waist >= 24.5 && waist <= 26 && hips >= 34.5 && hips <= 36) {
                    size4Count++;
                  }
                  else if (chest >= 34.5 && chest <= 36 && waist >= 26.5 && waist <= 27 && hips >= 36.5 && hips <= 37) {
                    size6Count++;
                  }
                  else if (chest >= 35.5 && chest <= 37 && waist >= 27.5 && waist <= 28 && hips >= 37.5 && hips <= 38) {
                    size8Count++;
                  }
                  else if (chest >= 36.5 && chest <= 39 && waist >= 28.5 && waist <= 30 && hips >= 38.5 && hips <= 40) {
                    size10Count++;
                  }
                  else if (chest >= 38.5 && chest <= 40 && waist >= 30.5 && waist <= 32 && hips >= 40.5 && hips <= 42) {
                    size12Count++;
                  }
                  else if (chest >= 39.5 && chest <= 41 && waist >= 32.5 && waist <= 34 && hips >= 42.5 && hips <= 43) {
                    size14Count++;
                  }
                  else if (chest >= 40.5 && chest <= 43 && waist >= 34.5 && waist <= 35 && hips >= 43.5 && hips <= 44) {
                    size16Count++;
                  }
                  else if (chest >= 41.5 && chest <= 44 && waist >= 35.5 && waist <= 37 && hips >= 44.5 && hips <= 46) {
                    size18Count++;
                  }
                  else if (chest >= 43.5 && chest <= 46 && waist >= 37.5 && waist <= 39 && hips >= 46.5 && hips <= 48) {
                    size20Count++;
                  }
                  else if (chest >= 45.5 && chest <= 48 && waist >= 39.5 && waist <= 41 && hips >= 48.5 && hips <= 50) {
                    size22Count++;
                  }
                  else if (chest >= 47.5 && chest <= 50 && waist >= 41.5 && waist <= 43 && hips >= 50.5 && hips <= 52) {
                    size24Count++;
                  }
                  else if (chest >= 49.5 && chest <= 52 && waist >= 43.5 && waist <= 45 && hips >= 52.5 && hips <= 54) {
                    size26Count++;
                  }
                  else if (chest >= 51.5 && chest <= 54 && waist >= 45.5 && waist <= 47 && hips >= 54.5 && hips <= 56) {
                    size28Count++;
                }
            }
        }
    }
    sheet.getRange("G23").setValue(size0Count);
    sheet.getRange("G24").setValue(size2Count);
    sheet.getRange("G25").setValue(size4Count);
    sheet.getRange("G26").setValue(size6Count);
    sheet.getRange("G27").setValue(size8Count);
    sheet.getRange("G28").setValue(size10Count);
    sheet.getRange("G29").setValue(size12Count);
    sheet.getRange("G30").setValue(size14Count);
    sheet.getRange("G31").setValue(size16Count);
    sheet.getRange("G32").setValue(size18Count);
    sheet.getRange("G33").setValue(size20Count);
    sheet.getRange("G34").setValue(size22Count);
    sheet.getRange("G35").setValue(size24Count);
    sheet.getRange("G36").setValue(size26Count);
    sheet.getRange("G37").setValue(size28Count);
    
    //Tie Format
    sheet.getRange("h22:i22").merge().setValue("Tie Order").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("h22:i22").setBackground("#213483").setFontColor("#FFFFFF");
    sheet.getRange("h22:i22").setBorder(true, true, true, true, true, true);
    sheet.getRange("h23:i23").merge().setBorder(true, true, true, true, true, true).setHorizontalAlignment("center").setFontWeight("bold");

    //tie order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var f9value = currentSheet.getRange("H10").getValue();
            if (f9value === true) {
                tieOrderCount++;
            }
        }
    }

}

