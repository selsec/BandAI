function setSectionColors() {
    var fluteColor = "Yellow";
    var clarinetColor = "Red";
    var saxophoneColor = "Blue";
    var trumpetColor = "White";
    var colorguardColor = "Pink";
    var mellophoneColor = "Orange";
    var lowBrassColor = "Teal";
    var tubaColor = "Purple";
    var percussionColor = "Green";

    // Prompt user for color input
    var fluteInput = prompt("Enter color for flute section (default: Yellow):");
    var clarinetInput = prompt("Enter color for clarinet section (default: Red):");
    var saxophoneInput = prompt("Enter color for saxophone section (default: Blue):");
    var trumpetInput = prompt("Enter color for trumpet section (default: White):");
    var colorguardInput = prompt("Enter color for colorguard section (default: Pink):");
    var mellophoneInput = prompt("Enter color for mellophone section (default: Orange):");
    var lowBrassInput = prompt("Enter color for low brass section (default: Teal):");
    var tubaInput = prompt("Enter color for tuba section (default: Purple):");
    var percussionInput = prompt("Enter color for percussion section (default: Green):");

    // Update section colors if input is provided
    if (fluteInput) {
        fluteColor = fluteInput;
    }
    if (clarinetInput) {
        clarinetColor = clarinetInput;
    }
    if (saxophoneInput) {
        saxophoneColor = saxophoneInput;
    }
    if (trumpetInput) {
        trumpetColor = trumpetInput;
    }
    if (colorguardInput) {
        colorguardColor = colorguardInput;
    }
    if (mellophoneInput) {
        mellophoneColor = mellophoneInput;
    }
    if (lowBrassInput) {
        lowBrassColor = lowBrassInput;
    }
    if (tubaInput) {
        tubaColor = tubaInput;
    }
    if (percussionInput) {
        percussionColor = percussionInput;
    }

    //getters for section colors
    function getFluteColor() {
        return fluteColor;
    }

    function getClarinetColor() {
        return clarinetColor;
    }

    function getSaxophoneColor() {
        return saxophoneColor;
    }

    function getTrumpetColor() {
        return trumpetColor;
    }

    function getColorguardColor() {
        return colorguardColor;
    }

    function getMellophoneColor() {
        return mellophoneColor;
    }

    function getLowBrassColor() {
        return lowBrassColor;
    }

    function getTubaColor() {
        return tubaColor;
    }

    function getPercussionColor() {
        return percussionColor;
    }
}