  //function to prompt user for color input and update the colors
  function setSectionColors() {
    var ui = SpreadsheetApp.getUi();
    
    //prompt user for color input and update the colors if input is provided
    var fluteInput = ui.prompt("Enter color for flute section (default: Yellow):").getResponseText();
    if (fluteInput) fluteColor = fluteInput;
  
    var clarinetInput = ui.prompt("Enter color for clarinet section (default: Red):").getResponseText();
    if (clarinetInput) clarinetColor = clarinetInput;
  
    var saxophoneInput = ui.prompt("Enter color for saxophone section (default: Blue):").getResponseText();
    if (saxophoneInput) saxophoneColor = saxophoneInput;
  
    var trumpetInput = ui.prompt("Enter color for trumpet section (default: White):").getResponseText();
    if (trumpetInput) trumpetColor = trumpetInput;
  
    var colorguardInput = ui.prompt("Enter color for colorguard section (default: Pink):").getResponseText();
    if (colorguardInput) colorguardColor = colorguardInput;
  
    var mellophoneInput = ui.prompt("Enter color for mellophone section (default: Orange):").getResponseText();
    if (mellophoneInput) mellophoneColor = mellophoneInput;
  
    var lowBrassInput = ui.prompt("Enter color for low brass section (default: Teal):").getResponseText();
    if (lowBrassInput) lowBrassColor = lowBrassInput;
  
    var tubaInput = ui.prompt("Enter color for tuba section (default: Purple):").getResponseText();
    if (tubaInput) tubaColor = tubaInput;
  
    var percussionInput = ui.prompt("Enter color for percussion section (default: Green):").getResponseText();
    if (percussionInput) percussionColor = percussionInput;
  }
  
  //getter functions to access the colors
  function getFluteColor() {
    return sectionColors.fluteColor;
  }
  
  function getClarinetColor() {
    return sectionColors.clarinetColor;
  }
  
  function getSaxophoneColor() {
    return sectionColors.saxophoneColor;
  }
  
  function getTrumpetColor() {
    return sectionColors.trumpetColor;
  }
  
  function getColorguardColor() {
    return sectionColors.colorguardColor;
  }
  
  function getMellophoneColor() {
    return sectionColors.mellophoneColor;
  }
  
  function getLowBrassColor() {
    return sectionColors.lowBrassColor;
  }
  
  function getTubaColor() {
    return sectionColors.tubaColor;
  }
  
  function getPercussionColor() {
    return sectionColors.percussionColor;
  }

  