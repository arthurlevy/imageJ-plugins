// This is a macro for automation of void content measurement in a micrograph.
// It works as a plugin in imageJ software.
// This version is for batch processing every file in a directory.
// Arthur LEVY 7-26-11

dir = getDirectory("Choose the directory with image files to process ");

list = getFileList(dir);
File.makeDirectory(dir+"outputs");

//void_content.xls headers.
File.append("file name\tvoid content\tlower threshold\tupper threshold\tMaximum Void Area (um^2)\tAverage Void Area (um^2)", dir+"outputs"+File.separator+"void_content.xls")

//Set options
run("Set Measurements...", "area mean center fit shape area_fraction redirect=None decimal=3");
    
// every file in the directory:
for (i=0; i<list.length; i++)
{
	path = dir+list[i];
    showProgress(i, list.length);

    //check if it's an image
    extensions = newArray("tif", "tiff", "jpg", "bmp");
    isImage = false;

    for (jext=0; jext<extensions.length; jext++)
    {
    	if (endsWith(toLowerCase(list[i]), "." + extensions[jext]))
        {
            isImage = true;
        }
	}
  
    if  (isImage)
    {
    	open(path);
        filename=File.nameWithoutExtension;
        
        //convert to grayscale
        run("8-bit");
        getDimensions(imwidth, imheight, channels, slices, frames);
        
        
        //setting scale
        mag =2;
        Dialog.create("Set Magnification");
        Dialog.addMessage("Treating file "+filename+"\nWhat is the microscope magnification\nused for this micrograph\?");
        Dialog.addNumber("X", 20);
        Dialog.addCheckbox("Skip this file.", false)
        Dialog.show();
        mag = Dialog.getNumber()/10;
        skip = Dialog.getCheckbox();
        if (!skip)
        {
                
        	//open file for output with prior checking for existing file
            if (File.exists(dir+"outputs"+File.separator+filename+".xls"))
            {
            	if (getBoolean("file "+filename+".xls exists. Overwrite\?"))
            	{
            		output_file = dir+"outputs"+File.separator+filename+".xls";
                    d=File.open(output_file);
                }
                else
                {
                	d=File.open("");
                	output_file = "user defined place";
                }
            }
                        
            else
            {
            	output_file = dir+"outputs"+File.separator+filename+".xls";
            	d=File.open(output_file);
            }
                        
                        
             //confirming scale
             scale_opt = "distance="+mag+" known=1 pixel=1 unit=um" ;
             run("Set Scale...", scale_opt);
             makeOval(imwidth/2, imheight/2, 7*mag, 7*mag);
             run("Clear Results");
                        
             //rotate
             setTool("line");
             waitForUser("Current selection should be the typical size of a carbon fiber.\n \nNow, draw an horizontal line for fitting rotation.");
             if (selectionType()==5)
             {
             	run("Measure");
             	angle = getResult("Angle");
             	selectWindow("Results");
             	run("Close");
             }
             else angle = 0;
             run("Rotate... ", "angle="+angle+" grid=1 interpolation=Bilinear");
             print(d,"Rotation Angle (deg)");
             print(d,angle);
             run("Select None");
             
             //Crop
             setTool("rectangle");
             waitForUser("Select the region of interest and click OK.");
             if (selectionType()!=0)
             {
             	run("Select All");
             }
                                    
             getSelectionBounds(x, y, width, height);
             print(d,"Region of interest");
             print(d,"xmin\t"+x);
             print(d,"width\t"+width);
             print(d,"ymin\t"+y);
             print(d,"heigth\t"+height);
             run("Crop");
             saveAs("jpeg", dir+"outputs"+File.separator+filename+"_crop.jpg");
             print(d," ");
             
             
             //get histogram before thresholding
             getHistogram(values, grayscale_hist, 256);
             //Threshold
             setThreshold(0, 200);
             run("Threshold...");
             selectWindow("Threshold");
             setLocation(screenWidth-300, screenHeight-250);
             waitForUser("Thresholding done \?\nBe sure to press the \"Set\" button and confirm the values.");
             getThreshold(lower, upper);
             print(d, "thresholds");
             print(d, lower+"\t"+upper);
             print(d," ");
             //convert to black and white
             run("Convert to Mask");
             
             //we don't fill holes cause they may be fibers
             //run("Fill Holes");
             
             //Void content
             run("Select All");
             getStatistics(area, mean, min, max, std, histogram);
             print(d, "Void Content [%]");
             print(d, mean/2.55);
             print(d," ");

             //through x distribution
             run("Select All");
             run("Clear Results");
             profile = getProfile();
             print(d, "Horizontal Void content distribution");
             print(d, "Position [um]\tvoid content [%]");
             for (k=0; k<profile.length; k++)
             {
             	print(d, k/mag+"\t"+profile[k]/2.55);
             }
             print(d," ");
             
             
             //through thickness distribution
             run("Select All");
             run("Clear Results");
             setKeyDown("alt"); //for vertical distribution
             profile = getProfile();
             print(d, "Vertical Void content distribution");
             print(d, "Position [um]\tvoid content [%]");
             for (k=0; k<profile.length; k++)
             {
             	print(d, k/mag+"\t"+profile[k]/2.55);
             }
             print(d," ");
             
             //Void sizes analysis
             run("Analyze Particles...", "size=0-Infinity circularity=0.00-1.00 show=Nothing");
             print(d,"Void size analysis");
             print(d,"X\tY\tArea(um^2)\tfitted ellipse");
             print(d,"\t\t\tMajor axis\tMinor axis\tangle");
             array_particle_size = newArray(nResults);
             for (k=0; k<nResults; k++)
             {
             	array_particle_size[k] = getResult("Area", k);
             }
             Array.getStatistics(array_particle_size, minVoidArea, maxVoidArea, meanVoidArea, stdDev);
             //append void content to void_content.xls
             File.append(filename+"\t"+mean/2.55+"\t"+lower+"\t"+upper+"\t"+maxVoidArea+"\t"+meanVoidArea, dir+"outputs"+File.separator+"void_content.xls")
             
             
             for (k=0; k<nResults; k++) 
             {
             	print(d, getResult("XM", k)+"\t"+getResult("YM", k)+"\t"+getResult("Area", k)+"\t"+getResult("Major", k)+"\t"+getResult("Minor", k)+"\t"+getResult("Angle", k));
             }
             print(d," ");
             
             // store the histogram values (computed in the firt place)
             print(d, "Intensity Histogram");
             print(d, "Grayscale Value\tnumber of pixels");
             for (k=0; k<256; k++)
             {
             	print(d, k+"\t"+grayscale_hist[k]);
             }
             print(d," ");
             
             File.close(d); 
             
             //cleaning
             selectWindow("Threshold");
             run("Close");
         }//if not skiped
         
         run("Close All");
     }// if isImage
     
 }//for eqch file
 
 
 //Done message
 showMessage("Processing completed", "you're done, data are stored in file\n    void_content.xls\nand every image has its cropped file and results excel file associated.\nAll files are in folder\n    "+dir+"outputs");
                                                    
