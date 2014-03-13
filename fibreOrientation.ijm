//this pluggin is to measure fibre orientation in a composite based on a micrograph


run("8-bit");
setAutoThreshold("Default");

//Threshold
setThreshold(0, 200);
run("Threshold...");
selectWindow("Threshold");
setLocation(screenWidth-300, screenHeight-250);
waitForUser("Thresholding done \?\nBe sure to press the \"Set\" button and confirm the values.");
getThreshold(lower, upper);
run("Convert to Mask");

run("Fill Holes");
run("Watershed");
run("Set Measurements...", "  fit redirect=None decimal=3");
run("Analyze Particles...", "size=500-5000 circularity=0.00-1.00 show=Ellipses display exclude in_situ");
saveAs("Results", "./Results.xls");
