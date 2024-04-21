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
        var smallCount = 0;
        var mediumCount = 0;
        var largeCount = 0;
        var xLargeCount = 0;
        var xxLargeCount = 0;
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Period Roster" && sheetName !== "Attendance" && sheetName !== "Uniform Order") {
            var a8Value = currentSheet.getRange("C8").getValue();
            if (a8Value === "S") {
                smallCount++;
            } else if (a8Value === "M") {
                mediumCount++;
            } else if (a8Value === "L") {
                largeCount++;
            } else if (a8Value === "XL") {
                xLargeCount++;
            } else if (a8Value === "XXL") {
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
    sheet.getRange("D2:H2").merge().setValue("Section Shirt Order").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("D2:H2").setBackground("#213483").setFontColor("#FFFFFF");
    sheet.getRange("D2:H2").setBorder(true, true, true, true, true, true);
    sheet.getRange("D3").setValue("Flute").setFontWeight("bold");
    sheet.getRange("E3").setValue(uniformOrder.getFluteColor()).setFontWeight("bold");
    sheet.getRange("D4").setValue("S").setFontWeight("bold");
    sheet.getRange("D5").setValue("M").setFontWeight("bold");
    sheet.getRange("D6").setValue("L").setFontWeight("bold");
    sheet.getRange("D7").setValue("XL").setFontWeight("bold");
    sheet.getRange("D8").setValue("XXL").setFontWeight("bold");
    sheet.getRange("D9").setValue("Clarinet").setFontWeight("bold");
    sheet.getRange("E9").setValue(uniformOrder.getClarinetColor()).setFontWeight("bold");
    sheet.getRange("D10").setValue("S").setFontWeight("bold");
    sheet.getRange("D11").setValue("M").setFontWeight("bold");
    sheet.getRange("D12").setValue("L").setFontWeight("bold");
    sheet.getRange("D13").setValue("XL").setFontWeight("bold");
    sheet.getRange("D14").setValue("XXL").setFontWeight("bold");
    sheet.getRange("D15").setValue("Saxophone").setFontWeight("bold");
    sheet.getRange("E15").setValue(uniformOrder.getSaxophoneColor()).setFontWeight("bold");
    sheet.getRange("D16").setValue("S").setFontWeight("bold");
    sheet.getRange("D17").setValue("M").setFontWeight("bold");
    sheet.getRange("D18").setValue("L").setFontWeight("bold");
    sheet.getRange("D19").setValue("XL").setFontWeight("bold");
    sheet.getRange("D20").setValue("XXL").setFontWeight("bold");
    sheet.getRange("F3").setValue("Trumpet").setFontWeight("bold");
    sheet.getRange("G3").setValue(uniformOrder.getTrumpetColor()).setFontWeight("bold");
    sheet.getRange("F4").setValue("S").setFontWeight("bold");
    sheet.getRange("F5").setValue("M").setFontWeight("bold");
    sheet.getRange("F6").setValue("L").setFontWeight("bold");
    sheet.getRange("F7").setValue("XL").setFontWeight("bold");
    sheet.getRange("F8").setValue("XXL").setFontWeight("bold");
    sheet.getRange("F9").setValue("Mellophone").setFontWeight("bold");
    sheet.getRange("G9").setValue(uniformOrder.getMellophoneColor()).setFontWeight("bold");
    sheet.getRange("F10").setValue("S").setFontWeight("bold");
    sheet.getRange("F11").setValue("M").setFontWeight("bold");
    sheet.getRange("F12").setValue("L").setFontWeight("bold");
    sheet.getRange("F13").setValue("XL").setFontWeight("bold");
    sheet.getRange("F14").setValue("XXL").setFontWeight("bold");
    sheet.getRange("F15").setValue("Low Brass").setFontWeight("bold");
    sheet.getRange("G15").setValue(uniformOrder.getLowBrassColor()).setFontWeight("bold");
    sheet.getRange("F16").setValue("S").setFontWeight("bold");
    sheet.getRange("F17").setValue("M").setFontWeight("bold");
    sheet.getRange("F18").setValue("L").setFontWeight("bold");
    sheet.getRange("F19").setValue("XL").setFontWeight("bold");
    sheet.getRange("F20").setValue("XXL").setFontWeight("bold");
    sheet.getRange("H3").setValue("Tuba").setFontWeight("bold");
    sheet.getRange("I3").setValue(uniformOrder.getTubaColor()).setFontWeight("bold");
    sheet.getRange("H4").setValue("S").setFontWeight("bold");
    sheet.getRange("H5").setValue("M").setFontWeight("bold");
    sheet.getRange("H6").setValue("L").setFontWeight("bold");
    sheet.getRange("H7").setValue("XL").setFontWeight("bold");
    sheet.getRange("H8").setValue("XXL").setFontWeight("bold");
    sheet.getRange("H9").setValue("Percussion").setFontWeight("bold");
    sheet.getRange("I9").setValue(uniformOrder.getPercussionColor()).setFontWeight("bold");
    sheet.getRange("H10").setValue("S").setFontWeight("bold");
    sheet.getRange("H11").setValue("M").setFontWeight("bold");
    sheet.getRange("H12").setValue("L").setFontWeight("bold");
    sheet.getRange("H13").setValue("XL").setFontWeight("bold");
    sheet.getRange("H14").setValue("XXL").setFontWeight("bold");
    sheet.getRange("H15").setValue("Colorguard").setFontWeight("bold");
    sheet.getRange("I15").setValue(uniformOrder.getColorguardColor()).setFontWeight("bold");
    sheet.getRange("H16").setValue("S").setFontWeight("bold");
    sheet.getRange("H17").setValue("M").setFontWeight("bold");
    sheet.getRange("H18").setValue("L").setFontWeight("bold");
    sheet.getRange("H19").setValue("XL").setFontWeight("bold");
    sheet.getRange("H20").setValue("XXL").setFontWeight("bold");
    sheet.getRange("D3:H20").setBorder(true, true, true, true, true, true);

    //fill data for section shirt order
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
    //flute section order
    var sheets = ss.getSheets();
    var fluteSmallCount = 0;
    var fluteMediumCount = 0;
    var fluteLargeCount = 0;
    var fluteXLargeCount = 0;
    var fluteXXLargeCount = 0;
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e9Value = currentSheet.getRange("E9").getValue();
            var a8Value = currentSheet.getRange("C8").getValue();
            if (e9Value === "flute" && a8Value === "S") {
                fluteSmallCount++;
            } else if (e9Value === "flute" && a8Value === "M") {
                fluteMediumCount++;
            } else if (e9Value === "flute" && a8Value === "L") {
                fluteLargeCount++;
            } else if (e9Value === "flute" && a8Value === "XL") {
                fluteXLargeCount++;
            } else if (e9Value === "flute" && a8Value === "XXL") {
                fluteXXLargeCount++;
            }
        }
    }
    sheet.getRange("E3").setValue(fluteSmallCount);
    sheet.getRange("E4").setValue(fluteMediumCount);
    sheet.getRange("E5").setValue(fluteLargeCount);
    sheet.getRange("E6").setValue(fluteXLargeCount);
    sheet.getRange("E7").setValue(fluteXXLargeCount);
    //clarinet section order
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var e9Value = currentSheet.getRange("E9").getValue();
            var a8Value = currentSheet.getRange("C8").getValue();
            if (e9Value === "clarinet" && a8Value === "S") {
                clarinetSmallCount++;
            } else if (e9Value === "clarinet" && a8Value === "M") {
                clarinetMediumCount++;
            } else if (e9Value === "clarinet" && a8Value === "L") {
                clarinetLargeCount++;
            } else if (e9Value === "clarinet" && a8Value === "XL") {
                clarinetXLargeCount++;
            } else if (e9Value === "clarinet" && a8Value === "XXL") {
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
            var e9Value = currentSheet.getRange("E9").getValue();
            var a8Value = currentSheet.getRange("C8").getValue();
            if (e9Value === "saxophone" && a8Value === "S") {
                saxophoneSmallCount++;
            } else if (e9Value === "saxophone" && a8Value === "M") {
                saxophoneMediumCount++;
            } else if (e9Value === "saxophone" && a8Value === "L") {
                saxophoneLargeCount++;
            } else if (e9Value === "saxophone" && a8Value === "XL") {
                saxophoneXLargeCount++;
            } else if (e9Value === "saxophone" && a8Value === "XXL") {
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
            var e9Value = currentSheet.getRange("E9").getValue();
            var a8Value = currentSheet.getRange("C8").getValue();
            if (e9Value === "trumpet" && a8Value === "S") {
                trumpetSmallCount++;
            } else if (e9Value === "trumpet" && a8Value === "M") {
                trumpetMediumCount++;
            } else if (e9Value === "trumpet" && a8Value === "L") {
                trumpetLargeCount++;
            } else if (e9Value === "trumpet" && a8Value === "XL") {
                trumpetXLargeCount++;
            } else if (e9Value === "trumpet" && a8Value === "XXL") {
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
            var e9Value = currentSheet.getRange("E9").getValue();
            var a8Value = currentSheet.getRange("C8").getValue();
            if (e9Value === "mellophone" && a8Value === "S") {
                mellophoneSmallCount++;
            } else if (e9Value === "mellophone" && a8Value === "M") {
                mellophoneMediumCount++;
            } else if (e9Value === "mellophone" && a8Value === "L") {
                mellophoneLargeCount++;
            } else if (e9Value === "mellophone" && a8Value === "XL") {
                mellophoneXLargeCount++;
            } else if (e9Value === "mellophone" && a8Value === "XXL") {
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
            var e9Value = currentSheet.getRange("E9").getValue();
            var a8Value = currentSheet.getRange("C8").getValue();
            if ((e9Value === "trombone" || e9Value === "baritone") && a8Value === "S") {
                lowBrassSmallCount++;
            } else if ((e9Value === "trombone" || e9Value === "baritone") && a8Value === "M") {
                lowBrassMediumCount++;
            } else if ((e9Value === "trombone" || e9Value === "baritone") && a8Value === "L") {
                lowBrassLargeCount++;
            } else if ((e9Value === "trombone" || e9Value === "baritone") && a8Value === "XL") {
                lowBrassXLargeCount++;
            } else if ((e9Value === "trombone" || e9Value === "baritone") && a8Value === "XXL") {
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
            var e9Value = currentSheet.getRange("E9").getValue();
            var a8Value = currentSheet.getRange("C8").getValue();
            if (e9Value === "tuba" && a8Value === "S") {
                tubaSmallCount++;
            } else if (e9Value === "tuba" && a8Value === "M") {
                tubaMediumCount++;
            } else if (e9Value === "tuba" && a8Value === "L") {
                tubaLargeCount++;
            } else if (e9Value === "tuba" && a8Value === "XL") {
                tubaXLargeCount++;
            } else if (e9Value === "tuba" && a8Value === "XXL") {
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
            var e9Value = currentSheet.getRange("E9").getValue();
            var a8Value = currentSheet.getRange("C8").getValue();
            if (e9Value === "percussion" && a8Value === "S") {
                percussionSmallCount++;
            } else if (e9Value === "percussion" && a8Value === "M") {
                percussionMediumCount++;
            } else if (e9Value === "percussion" && a8Value === "L") {
                percussionLargeCount++;
            } else if (e9Value === "percussion" && a8Value === "XL") {
                percussionXLargeCount++;
            } else if (e9Value === "percussion" && a8Value === "XXL") {
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
            var e9Value = currentSheet.getRange("E9").getValue();
            var a8Value = currentSheet.getRange("C8").getValue();
            if (e9Value === "colorguard" && a8Value === "S") {
                colorguardSmallCount++;
            } else if (e9Value === "colorguard" && a8Value === "M") {
                colorguardMediumCount++;
            } else if (e9Value === "colorguard" && a8Value === "L") {
                colorguardLargeCount++;
            } else if (e9Value === "colorguard" && a8Value === "XL") {
                colorguardXLargeCount++;
            } else if (e9Value === "colorguard" && a8Value === "XXL") {
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
    
    
    for (var i = 0; i < sheets.length; i++) {
        var currentSheet = sheets[i];
        var sheetName = currentSheet.getName();
        if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Uniform Order" && sheetName !== "Dashboard" && sheetName !== "IncomeExpense") {
            var a9Value = currentSheet.getRange("A9").getValue();
            if (a9Value === true) {
                var a8Value = currentSheet.getRange("A8").getValue();
                if (a8Value === "5.5") {
                    marchingShoes55Count++;
                } else if (a8Value === "6") {
                    marchingShoes6Count++;
                }
                else if (a8Value === "6.5") {
                    marchingShoes65Count++;
                }
                else if (a8Value === "7") {
                    marchingShoes7Count++;
                }
                else if (a8Value === "7.5") {
                    marchingShoes75Count++;
                }
                else if (a8Value === "8") {
                    marchingShoes8Count++;
                }
                else if (a8Value === "8.5") {
                    marchingShoes85Count++;
                }
                else if (a8Value === "9") {
                    marchingShoes9Count++;
                }
                else if (a8Value === "9.5") {
                    marchingShoes95Count++;
                }
                else if (a8Value === "10") {
                    marchingShoes10Count++;
                }
                else if (a8Value === "10.5") {
                    marchingShoes105Count++;
                }
                else if (a8Value === "11") {
                    marchingShoes11Count++;
                }
                else if (a8Value === "11.5") {
                    marchingShoes115Count++;
                }
                else if (a8Value === "12") {
                    marchingShoes12Count++;
                }
                else if (a8Value === "12.5") {
                    marchingShoes125Count++;
                }
                else if (a8Value === "13") {
                    marchingShoes13Count++;
                }
                else if (a8Value === "13.5") {
                    marchingShoes135Count++;
                }
                else if (a8Value === "14") {
                    marchingShoes14Count++;
                }
                else {
                    marchingShoesOtherCount = marchingShoesOtherCount + ", " + a8Value;
                }
            }
        }
    }
        //bibbers order format
}