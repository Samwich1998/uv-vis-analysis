
"""
Need to Install in the Python Enviroment Beforehand:
    $ conda install openpyxl
    $ conda install scipy
    $ pip install natsort
    $ pip install pyexcel
    $ pip install BaselineRemoval
"""

# -------------------------------------------------------------------------- #
# ---------------------------- Imported Modules ---------------------------- #

# Import Basic Modules
import os
import sys
import numpy as np
from pathlib import Path
# Module to Sort Files in Order
from natsort import natsorted
# Modules to Plot
import matplotlib.pyplot as plt
# Modules for Filtering
from scipy.signal import savgol_filter

# Import Python Helper Files
sys.path.append('./Helper Files/')  # Folder with All the Helper Files
import excelProcessing
import calculateBaseline

# ---------------------------------------------------------------------------#
# ---------------------------------------------------------------------------#


if __name__ == "__main__":
    # ---------------------------------------------------------------------- #
    #    User Parameters to Edit (More Complex Edits are Inside the Files)   #
    # ---------------------------------------------------------------------- #

    # Specify Where the Files are Located
    dataDirectory = "./data/05-20-2022 DA MIP 5 Hours/"   # The Path to the Data; Must End With '/'
    
    # Specify Which Files to Read In
    useAllFolderFiles = True # Read in All TXT/CSV/EXCEL Files in the dataDirectory
    if useAllFolderFiles:
        # Specify Which Files You Want to Read
        fileDoesntContain = "N/A"
        fileContains = ""
    else:
        # Else, Specify the File Names
        dpvFiles = ['UV-Vis 4_27_2022 3_56_28 PM.tsv']

    # Specify the Plotting Extent
    plotBaselineSteps = False # Display the Baseline as Well as the Final Current After Baseline Subtraction
    
    # Apply bounds for peak detection
    applyBoundsForPeak = True
    peakWavelengthBounds = [240, 320]
    
    # Are we Looking for a Positive Peak?
    isPeakPositive = True
    
    # ---------------------------------------------------------------------- #
    # ------------------------- Preparation Steps -------------------------- #
    
    # Define the Baseline Remover Class
    baselineObject = calculateBaseline.bestLinearFit()
    
    # Get File Information
    extractData = excelProcessing.processFiles()
    if useAllFolderFiles:
        dpvFiles = extractData.getFiles(dataDirectory, fileDoesntContain, fileContains)
    # Sort Files
    dpvFiles = natsorted(dpvFiles)
    
    # ---------------------------------------------------------------------- #
    # ----------------------------- DPV Program ---------------------------- #
    
    # For Each File, Extract the Important Data ,and Plot
    for fileNum, currentFile in enumerate(sorted(dpvFiles)):
        # Create Output Folder if the One Given Does Not Exist
        outputDirectory = dataDirectory + "Analysis/" + Path(currentFile).stem + "/"
        os.makedirs(outputDirectory, exist_ok = True)
        
        # ----------------------- Extract the Data --------------------------#
        # Extract the Data/File Information from the File (Potential, Current)
        dataFile = dataDirectory + currentFile
        fileName = os.path.splitext(currentFile)[0]
        wavelengthList, absorbanceList, sampleNames = extractData.getData(dataFile, outputDirectory, testSheetNum = 0)
        # ------------------------------------------------------------------ # 
        
        analyzedData = []
        for sampleNum in range(len(sampleNames)):
            # Extract the Current Run
            sampleName = sampleNames[sampleNum]
            wavelength = np.array(wavelengthList[sampleNum])
            absorbance = np.array(absorbanceList[sampleNum])
            
            if not isPeakPositive:
                absorbance = -absorbance
            # ----------------------- Filter the Data ---------------------- #
            # Low Pass Filter
            samplingFreq = len(wavelength)/(wavelength[-1] - wavelength[0])
            filteredData = baselineObject.butterFilter(absorbance, 0.1, samplingFreq, order = 3, filterType = 'low')
            # Apply a Savgol Filter
            filteredData = savgol_filter(filteredData, 15, 2, mode='nearest')
            # -------------------------------------------------------------- #
        
            # ------------------------ Find the Peak ----------------------- #
            # Find Peaks in the Data
            if applyBoundsForPeak:
                peakInd = baselineObject.findPeak(wavelength, filteredData, peakWavelengthBounds)
            else:
                peakInd = baselineObject.findPeak(wavelength, filteredData)
            # Return None if No Peak Found
            if peakInd == None:
                print("No Peak Found in " + sampleName + " Data")
                continue
            # -------------------------------------------------------------- #
    
            # ---------------- Find and Remove the Baseline ---------------- #
            # Get Baseline from Best Linear Fit
            leftCutInd, rightCutInd = baselineObject.findLinearBaseline(wavelength, filteredData, peakInd)
            if None in [leftCutInd, rightCutInd]:
                print("No Baseline Found in " + sampleName + " Data")
                continue
            
            # Fit Lines to Ends of Graph
            lineSlope, slopeIntercept = np.polyfit(wavelength[[leftCutInd, rightCutInd]], filteredData[[leftCutInd, rightCutInd]], 1)
            linearFit = lineSlope*wavelength + slopeIntercept

            # Piece Together absorbance's Baseline
            baseline = np.concatenate((filteredData[0:leftCutInd], linearFit[leftCutInd: rightCutInd+1], filteredData[rightCutInd+1:len(filteredData)]))
            # Find absorbance After Baseline Subtraction
            baselineData = filteredData - baseline
            
            # Get the Results
            peakInd = np.argmax(baselineData[leftCutInd:rightCutInd])
            peakWavelength = wavelength[leftCutInd:rightCutInd][peakInd]
            peakHeight = filteredData[leftCutInd:rightCutInd][peakInd]
            peakHeight_Baseline = baselineData[leftCutInd:rightCutInd][peakInd]
            analyzedData.append([sampleName, peakWavelength, peakHeight, peakHeight_Baseline])
            
            print(sampleName + ": Peak at (" + str(peakWavelength) + ", " + str(np.round(peakHeight_Baseline, 4)) + ")")
            # -------------------------------------------------------------- #
            
            # ------------------ Plot and Save the Results ----------------- #
            # Create a new figure
            figure = plt.figure()
            # Plot the Data
            plt.plot(wavelength, absorbance, linewidth=2, label = "Raw Data")
            plt.plot(wavelength, filteredData, linewidth=2, label = "Filtered Data")
            plt.plot(wavelength, baselineData, linewidth=2, label = "Baseline Subtracted")
            plt.plot(wavelength[[leftCutInd,rightCutInd]], filteredData[[leftCutInd,rightCutInd]], linewidth=2)
            # Set the Labels
            plt.title(sampleName + ": Peak at (" + str(peakWavelength) + ", " + str(np.round(peakHeight_Baseline, 4)) + ")")
            plt.xlabel("Wavelength (nm)")
            plt.ylabel("Absorbance (AU)")
            # Add the Legend
            legend = plt.legend()
            # Set the Figure Boundaries
            plt.xlim(wavelength[leftCutInd] - 20, wavelength[rightCutInd ]+ 20)
            plt.ylim(min(min(filteredData[leftCutInd:rightCutInd]), 0)*1.1,max(peakHeight_Baseline, peakHeight)*1.1)
            # Save the Data
            plt.savefig(outputDirectory + sampleName + ".png", dpi=300, bbox_extra_artists=(legend,), bbox_inches='tight')
            plt.show()
            # -------------------------------------------------------------- #
        
        # Save the Data
        saveData = excelProcessing.saveData()
        saveData.saveData(analyzedData, outputDirectory, fileName + " Analysis.xlsx")
            


