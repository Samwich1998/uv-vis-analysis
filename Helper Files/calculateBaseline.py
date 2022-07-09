
"""
Need to Install in the Python Enviroment Beforehand:
    $ pip install BaselineRemoval
"""

# Import Basic Modules
import numpy as np
# Import Modules for Filtering
from scipy.signal import butter
from scipy.signal import savgol_filter
# Import Modules to Find Peak
import scipy.signal
# Modules to Plot
import matplotlib.pyplot as plt


# ---------------------------------------------------------------------------#
# ---------------------------------------------------------------------------#

class bestLinearFit:
    
    def __init__(self):
        self.minLeftBoundaryInd = 50
        self.minPeakDuration = 10
        
    def butterParams(self, cutoffFreq = [0.1, 7], samplingFreq = 800, order=3, filterType = 'band'):
        nyq = 0.5 * samplingFreq
        if filterType == "band":
            normal_cutoff = [freq/nyq for freq in cutoffFreq]
        else:
            normal_cutoff = cutoffFreq / nyq
        sos = butter(order, normal_cutoff, btype = filterType, analog = False, output='sos')
        return sos
    
    def butterFilter(self, data, cutoffFreq, samplingFreq, order = 3, filterType = 'band'):
        sos = self.butterParams(cutoffFreq, samplingFreq, order, filterType)
        return scipy.signal.sosfiltfilt(sos, data)
    
    def findPeak(self, xData, yData, peakWavelengthBounds = [10, -10], deriv = False):
        # Find All Peaks in the Data
        peakInfo = scipy.signal.find_peaks(yData, prominence=10E-10, width=5, distance = 20)
        # Extract the Peak Information
        allProminences = peakInfo[1]['prominences']
        peakIndices = peakInfo[0]
        
        # Remove Peaks Nearby Boundaries
        allProminences = allProminences[np.logical_and(peakWavelengthBounds[0] < xData[peakIndices], xData[peakIndices] < peakWavelengthBounds[1])]
        peakIndices = peakIndices[np.logical_and(peakWavelengthBounds[0] < xData[peakIndices], xData[peakIndices] < peakWavelengthBounds[1])]
        # Seperate Out the Stimulus Window
        allProminences = allProminences[self.minLeftBoundaryInd < peakIndices]
        peakIndices = peakIndices[self.minLeftBoundaryInd < peakIndices]

        # If Peaks are Found
        if len(peakIndices) > 0:
            # Take the Most Prominent Peak
            bestPeak = allProminences.argmax()
            peakInd = peakIndices[bestPeak]
            return peakInd
        elif not deriv:
            filteredVelocity = savgol_filter(np.gradient(yData), 251, 3)
            return self.findPeak(xData, filteredVelocity, deriv = True)
        # If No Peak is Found, Return None
        return None
        
    
    def findLinearBaseline(self, xData, yData, peakInd):
        # Define a threshold for distinguishing good/bad lines
        maxBadPointsTotal = int(len(xData)/10)
        # Store Possibly Good Tangent Indexes
        goodTangentInd = [[] for _ in range(maxBadPointsTotal)]
        
        # For Each Index Pair on the Left and Right of the Peak
        for rightInd in range(peakInd+2, len(yData), 1):
            for leftInd in range(peakInd-2, self.minLeftBoundaryInd, -1):
                
                # Initialize range of data to check
                checkPeakBuffer = int((rightInd - leftInd)/4)
                xDataCut = xData[max(0, leftInd - checkPeakBuffer):rightInd + checkPeakBuffer]
                yDataCut = yData[max(0, leftInd - checkPeakBuffer):rightInd + checkPeakBuffer]
                
                # Draw a Linear Line Between the Points
                lineSlope = (yData[leftInd] - yData[rightInd])/(xData[leftInd] - xData[rightInd])
                slopeIntercept = yData[leftInd] - lineSlope*xData[leftInd]
                linearFit = lineSlope*xDataCut + slopeIntercept

                # Find the Number of Points Above the Tangent Line
                numWrongSideOfTangent = len(linearFit[linearFit - yDataCut > 0])

                # Define a threshold for distinguishing good/bad lines
                maxBadPoints = int(len(linearFit)/10) # Minimum 1/6
                if numWrongSideOfTangent < maxBadPoints and rightInd - leftInd > self.minPeakDuration:
                    goodTangentInd[numWrongSideOfTangent].append((leftInd, rightInd))
                    
        # If Nothing Found, Try and Return a Semi-Optimal Tangent Position
        for goodInd in range(maxBadPointsTotal):
            if len(goodTangentInd[goodInd]) != 0:
                return max(goodTangentInd[goodInd], key=lambda tangentPair: tangentPair[1]-tangentPair[0])
        return None, None
    
    def plotLinearFit(self, leftCutInd, rightCutInd, peakInd):
        plt.figure()
        plt.plot(self.potential, self.current);
        plt.plot(self.potential[[leftCutInd, rightCutInd, peakInd]],  self.current[[leftCutInd, rightCutInd, peakInd]], 'o');
        plt.plot(self.potential, self.linearFit, linewidth=0.3)
        plt.show()
        
        plt.figure()
        plt.plot(self.potential, self.current, label = "True Data")
        plt.plot(self.potential, self.baseline, label="Baseline Current")
        plt.show()
        


