# -*- coding: utf-8 -*-
"""
Created on Sat Feb 22 12:52:57 2020

@author: Mr. Toç
"""

import xlrd
from matplotlib import pyplot as plt
import xlsxwriter
from scipy.interpolate import CubicSpline
import numpy as np
from datetime import datetime

print('\n\n____Mr. Toç Code Industry is welcomed you to Soil Mechanics Test Analyzer____')

continuity = False
while continuity == False:
    try:
        test_wanted = int(input("\n\n1) Determination of Specific Gravity of Solids\n2) Grain Size Distribution by Sieve Analysis\n3) Grain Size Distribution by Sedimentation (Hydrometer)\n4) Atterberg (Consistency) Limits\n5) Standard Proctor Compaction Test\n\nIf you don have the Input Excel named 'SoilMechanicsDataSet.xlsx', please enter '7'.\n\nWhich test do you want to analyze?: "))
        continuity = True
    except ValueError:
        print('You Enter Wrong Input!!!')
        
temperature = ['2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', 
               '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31', '32', '33', '34', '35', '36', '37', '38', '39', '40']

density = ['0.9999', '1', '1', '1', '0.9999', '0.9999', '0.9999', '0.9998', '0.9997', '0.9996', '0.9995', '0.9994', 
           '0.9992', '0.9991', '0.9989', '0.9988', '0.9986', '0.9984', '0.9982', '0.998', '0.9978', '0.9975', 
           '0.9973', '0.997', '0.9968', '0.9965', '0.9962', '0.9959', '0.9956', '0.9953', '0.995', '0.9947', '0.9944', 
           '0.994', '0.9937', '0.9933', '0.993', '0.9926', '0.9922']

viscosity = ['1.6735', '1.619', '1.5673', '1.5182', '1.4715', '1.4271', '1.3847', '1.3444', '1.3059', '1.2692', '1.234', 
             '1.2005', '1.1683', '1.1375', '1.1081', '1.0798', '1.0526', '1.0266', '1.0016', '0.9775', '0.9544', '0.9321', 
             '0.9107', '0.89', '0.8701', '0.8509', '0.8324', '0.8145', '0.7972', '0.7805', '0.7644', '0.7488', '0.7337', 
             '0.7191', '0.705', '0.6913', '0.678', '0.6652', '0.6527']

if test_wanted == 7:
    excelldata = xlsxwriter.Workbook("SoilMechanicsDataSet.xlsx")
    sheetdata = excelldata.add_worksheet("Soil Mechanics Data Set")
    
    bold_format = excelldata.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})
    
    normal_format = excelldata.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})
    
    sheetdata.merge_range('A1:D1', 'Determination of Specific Gravity of Solids', bold_format)
    
    dataspec = ['Trial No', 'Mass of Empty Bottle (g)', 'Mass of Bottle Filled with Dry Soil (g)', 'Mass of Bottle with Soil, Filled with Deaired Water (g)', 'Temperature (˚C)', 'Mass of Bottle Filled with Water Only (g)', '1', '2', '3']
    
    for i in range(6):
        sheetdata.write(i+2, 0, dataspec[i], bold_format)
    for i in range(3):
        sheetdata.write_number(2, i+1, float(dataspec[i+6]), bold_format)
    
    datasieve = ['ASTM Designation', 'Nearest Size TS (mm)', 'Mass Retained (g)', 'No 3/4', 'No 3/8', 'No 4', 'No 10', 'No 30', 'No 40', 'No 50', 'No 100', 'No 200', '-']
    
    sheetdata.merge_range('G1:J1', 'Grain Size Distribution by Sieve Analysis', bold_format)
    sheetdata.merge_range('G3:H3', 'Test Sieve', bold_format)
    
    for i in range(10):
        sheetdata.write(i+4, 6, datasieve[i+3], bold_format)
    
    for i in range(3):
        sheetdata.write(3, i+6, datasieve[i], bold_format)
    
    sheetdata.write(15, 6, 'Total Mass of Dry Sample (g)', bold_format)
    sheetdata.write(13, 7, 'Pen', bold_format)
    
    datahydrometer = ['Total Dry Mass Tested (g)', 'Specific Gravity', 'Maximum Grain Size (mm)', 'Total Sample Percentage', 'Density of the Referenced Fluid (g/cc)', 
                      'Hydrometer Type', 'Volume of the Hydrometer (cm^3)', 'Area of the Cylinder (cm^2)', 'Meniscus Correction', 'Smaller Density, r1 (g/cc)', 'Maximum Density, r2 (g/cc)', 'Bigger, H1 (cm)', 'Smaller, H2 (cm)', 
                      'Time(min)', 'Temperature(Celcius)', 'Hydrometer Reading', '0.5', '1', '2', '4', '8', '15', '30', '60', '120', '240', '480', '960', '1440']
    
    sheetdata.merge_range('M1:O1', 'Grain Size Distribution by Sedimentation (Hydrometer)', bold_format)
    sheetdata.merge_range('M3:N3', 'Basic Properties', bold_format)
    sheetdata.merge_range('M10:N10', 'Hydrometer Properties', bold_format)
    sheetdata.merge_range('M20:O20', 'Hydrometer Readings', bold_format)
    
    for i in range(5):
        sheetdata.write(i+3, 12, datahydrometer[i], bold_format)
    
    for i in range(8):
        sheetdata.write(i+10, 12, datahydrometer[i+5], bold_format)
    
    for i in range(3):
        sheetdata.write(20, i+12, datahydrometer[i+13], bold_format)
    
    for i in range(13):
        sheetdata.write_number(i+21, 12, float(datahydrometer[i+16]), bold_format)
    
    
    dataatterberg = ['Type of Test', 'Number of Drops', 'Mass of Container + Wet Soil (g)', 'Mass of Container + Dry Soil (g)', 'Mass of Container (g)']
    
    sheetdata.merge_range('R1:W1', 'Atterberg (Consistency) Limits', bold_format)
    sheetdata.merge_range('S3:U3', 'Liquid Limit', bold_format)
    sheetdata.merge_range('V3:W3', 'Plastic Limit', bold_format)
    
    for i in range(5):
        sheetdata.write(i+2, 17, dataatterberg[i], bold_format)
    
    dataproctor = [['Total Mass of Sample (g)', 'Specific Gravity of the Sample', 'Coarser than 9.5 mm (%)', 'Water Content of Coarse (%)', 'Specific Gravity of the Removed Fraction', 'Height of the Mold (cm)', 'Diameter of Mold (cm)'], 
                    ['Test Number', 'Mass of Mold + Base (g)', 'Mass of Mold + Base + Compacted Soil (g)', '1', '2', '3', '4', '5'], 
                    ['Test Number', 'Container + Wet Sample (g)', 'Container + Dry Sample (g)', 'Mass of Container (g)']]
    
    sheetdata.merge_range('Z1:AE1', 'Standard Proctor Compaction Test', bold_format)
    sheetdata.merge_range('Z13:AE13', 'Moisture Content', bold_format)
    
    for i in range(5):
        sheetdata.write(8, i+26, float(dataproctor[1][i+3]), bold_format)
        sheetdata.write(13, i+26, float(dataproctor[1][i+3]), bold_format)
    for i in range(5):
        sheetdata.write(i+2, 25, dataproctor[0][i], bold_format)
    for i in range(2):
        sheetdata.write(i+2, 28, dataproctor[0][i+5], bold_format)
    for i in range(3):
        sheetdata.write(i+8, 25, dataproctor[1][i], bold_format)
    for i in range(4):
        sheetdata.write(i+13, 25, dataproctor[2][i], bold_format)
    
    excelldata.close()



"""-------------------SPECIFIC GRAVITY-------------------"""
if test_wanted == 1:
    #------------OPENING EXCEL--------------
    path = "SoilMechanicsDataSet.xlsx"
    databook = xlrd.open_workbook(path)
    datasheet = databook.sheet_by_index(0)
    



    calculateddata_specificgravity = [['Trial No', '1', '2', '3'], 
                                      ['Mass of Empty Bottle (g)'], 
                                      ['Mass of Bottle Filled with Dry Soil (g)'], 
                                      ['Mass of Bottle with Soil, Filled with Deaired Water (g)'], 
                                      ['Temperature (˚C)'], 
                                      ['Mass of Bottle Filled with Water Only (g)'], 
                                      ['Mass of Solids (g)'], 
                                      ['Mass of Water with Volume Equal to the Volume of Solids (g)'], 
                                      ['Specific Gravity of Solids at This Temperature'], 
                                      ['Specific Gravity of Solids at 20 ˚C'], 
                                      ['Mean Value of Specific Gravity'], 
                                      ['Standard Deviation of Specific Gravity']]
    
    def ms_calculator(mbs, mb):
        ms = mbs - mb
        return ms
    def equalm_calculator(mbw, mbs, mb, mbsw):
        equalm = mbw + mbs - mb - mbsw
        return equalm
    def realspec_calculator(ms, mbw, mbsw):
        realspec = ms / (mbw + ms - mbsw)
        return realspec
    def spec20_calculator(realspec, density, temperature, temp_spec):
        spec20 = (realspec * densFluid(density, temperature, temp_spec)) / 0.99821
        return spec20
    def densFluid(density, temperature, temp_spec):
            dens = float(density[temperature.index(str(int(temp_spec)))]) + ((float(temp_spec)-float(temperature[temperature.index(str(int(temp_spec)))]))/float(temperature[temperature.index(str(int(temp_spec))) + 1]) - float(temperature[temperature.index(str(int(temp_spec)))]))*(float(density[temperature.index(str(int(temp_spec))) + 1]) - float(density[temperature.index(str(int(temp_spec)))]))
            return dens
    
    #-------------CALCULATIONS-------------
    for i in range(3):
        mb = float(datasheet.cell_value(3, i+1))
        mbs = float(datasheet.cell_value(4, i+1))
        mbsw = float(datasheet.cell_value(5, i+1))
        temp_spec = float(datasheet.cell_value(6, i+1))
        mbw = float(datasheet.cell_value(7, i+1))
        
        ms = ms_calculator(mbs, mb)
        equalm = equalm_calculator(mbw, mbs, mb, mbsw)
        realspec = realspec_calculator(ms, mbw, mbsw)
        spec20 = spec20_calculator(realspec, density, temperature, temp_spec)
        
        calculateddata_specificgravity[1].append(str(mb))
        calculateddata_specificgravity[2].append(str(mbs))
        calculateddata_specificgravity[3].append(str(mbsw))
        calculateddata_specificgravity[4].append(str(temp_spec))
        calculateddata_specificgravity[5].append(str(mbw))
        calculateddata_specificgravity[6].append(str(ms))
        calculateddata_specificgravity[7].append(str(equalm))
        calculateddata_specificgravity[8].append(str(realspec))
        calculateddata_specificgravity[9].append(str(spec20))
    
    mean = (float(calculateddata_specificgravity[9][1]) + float(calculateddata_specificgravity[9][2]) + float(calculateddata_specificgravity[9][3])) / 3
    calculateddata_specificgravity[10].append(str(mean))
    
    stdv = ((float(calculateddata_specificgravity[9][1])-mean)**2 + (float(calculateddata_specificgravity[9][2])-mean)**2 + (float(calculateddata_specificgravity[9][3])-mean)**2)**0.5
    calculateddata_specificgravity[11].append(str(stdv))
    
    #-------------EXCEL PART-------------    
    excellspec = xlsxwriter.Workbook('Specific Gravity Data ' + str(datetime.now().year) + '.' + str(datetime.now().month) + '.' + str(datetime.now().day) + ' ' + str(datetime.now().hour) + '.' + str(datetime.now().minute) + '.' + str(datetime.now().second) + '.xlsx')
    specsheet = excellspec.add_worksheet('Specific Gravity Data Results')
    
    bold_format = excellspec.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})
    
    normal_format = excellspec.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})
    
    specsheet.merge_range('A1:D1', 'Specific Gravity Test Results', bold_format)
    
    for i in range(12):
        specsheet.write(i+1, 0, calculateddata_specificgravity[i][0], bold_format)
    
    for i in range(3):
        specsheet.write(1, i+1, calculateddata_specificgravity[0][i+1], bold_format)
    
    for i in range(9):
        for j in range(3):
            specsheet.write_number(i+2, j+1, float("%.4f" % float(calculateddata_specificgravity[i+1][j+1])), normal_format)
    
    specsheet.merge_range('B12:D12', "%.4f" % float(calculateddata_specificgravity[10][1]), normal_format)
    specsheet.merge_range('B13:D13', "%.4f" % float(calculateddata_specificgravity[11][1]), normal_format)
    
    excellspec.close()




"""-------------------SIEVE ANALYSIS-------------------"""
if test_wanted == 2:
    #------------OPENING EXCEL--------------
    path = "SoilMechanicsDataSet.xlsx"
    databook = xlrd.open_workbook(path)
    datasheet = databook.sheet_by_index(0)
    



    calculateddata_sieve = [['Test sieve'], 
                            ['ASTM Designation', 'Nearest TS Size (mm)', 'Mass Retained (g)', 'Cumulative Mass Retained (g)', 'Cumulative Retained (%)', 'Cumulative Passing'], 
                            ['No 3/4'], 
                            ['No 3/8'], 
                            ['No 4'], 
                            ['No 10'], 
                            ['No 30'], 
                            ['No 40'], 
                            ['No 50'], 
                            ['No 100'], 
                            ['No 200'], 
                            ['-', '0.074']]
    
    for i in range(9):
        nearest_size = float(datasheet.cell_value(i+4, 7))
        mass_retained = float(datasheet.cell_value(i+4, 8))
        calculateddata_sieve[i+2].append(str(nearest_size))
        calculateddata_sieve[i+2].append(str(mass_retained))
    calculateddata_sieve[11].append(str(float(datasheet.cell_value(13, 8))))
    calculateddata_sieve[2].append(calculateddata_sieve[2][2])
    for i in range(9):
        calculateddata_sieve[i+3].append(str(float(calculateddata_sieve[i+2][3]) + float(calculateddata_sieve[i+3][2])))
    
    totalmass = float(datasheet.cell_value(15, 7))
    
    for i in range(10):
        calculateddata_sieve[i+2].append(str((float(calculateddata_sieve[i+2][3])*100)/totalmass))
    
    for i in range(10):
        calculateddata_sieve[i+2].append(str(100 - float(calculateddata_sieve[i+2][4])))
    
    #---------------EXCELPART----------------
    sieveexcell = xlsxwriter.Workbook('Sieve Analysis Data ' + str(datetime.now().year) + '.' + str(datetime.now().month) + '.' + str(datetime.now().day) + ' ' + str(datetime.now().hour) + '.' + str(datetime.now().minute) + '.' + str(datetime.now().second) + '.xlsx')
    sievesheet = sieveexcell.add_worksheet('Sieve Analysis Data Results')
    
    bold_format = sieveexcell.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})
    
    normal_format = sieveexcell.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})
    
    sievesheet.merge_range('A1:F1', 'Sieve Analysis Test Results', bold_format)
    sievesheet.merge_range('A3:B3', 'Test Sieve', bold_format)
    
    for i in range(6):
        sievesheet.write(3, i, calculateddata_sieve[1][i], bold_format)
    for i in range(9):
        sievesheet.write(i+4, 0, calculateddata_sieve[i+2][0], bold_format)
    sievesheet.write(12, 1, calculateddata_sieve[11][1], bold_format)
    
    for i in range(10):
        for j in range(4):
            sievesheet.write_number(i+3, j+2, float("%.4f" % float(calculateddata_sieve[i+2][j+2])), normal_format)
    for i in range(9):
        sievesheet.write_number(i+3, 1, float("%.4f" % float(calculateddata_sieve[i+2][1])), normal_format)
    
    
    sieveexcell.close()
    
    
    x_sieve = []
    y_sieve = []
    for i in range(10):
        x_sieve.append(float(calculateddata_sieve[i+2][1]))
    for i in range(10):
        y_sieve.append(float(calculateddata_sieve[i+2][5]))
    x_sieve.reverse()
    y_sieve.reverse()
    
    spline = CubicSpline(x_sieve,y_sieve,bc_type='natural')
    
    x_spline_sieve = np.linspace(x_sieve[0], x_sieve[-1], 10000)
    y_spline_sieve = spline(x_spline_sieve)
    
    #-------------PLOTTING THE DISTRIBUTION TABLE------------
    plt.plot(x_spline_sieve, y_spline_sieve, color = 'b')
    
    plt.axhline(linewidth=2, color='k')
    
    
    plt.title('Sieve Analysis')
    plt.xlabel('Particle Size')
    plt.ylabel('%finer')
    plt.xscale('log', basex = 10)
    plt.grid(True, which="both")
    plt.gca().spines['top'].set_visible(False)
    plt.gca().spines['right'].set_visible(False)
    plt.gca().set_xlim([0.0001, 76.2])
    plt.gca().set_ylim([0, 103])
    
    plt.axvspan(0, 0.002, facecolor = '#AAD1D7')
    plt.axvspan(0.002, 0.074, facecolor = '#A2A785')
    plt.axvspan(0.074, 4.76, facecolor = '#8FFF64')
    plt.axvspan(4.76, 76.2, facecolor = '#E5ED27')
                
    plt.text(0.0005, 90, 'CLAY', horizontalalignment='center')
    plt.text(0.014, 90, 'SILT', horizontalalignment='center')
    plt.text(0.417, 90, 'SAND', horizontalalignment='center')
    plt.text(20.48, 90, 'GRAVEL', horizontalalignment='center')
    
    plt.show()




"""-------------------HYDROMETER TEST-------------------"""
if test_wanted == 3:
    #------------OPENING EXCEL--------------
    path = "SoilMechanicsDataSet.xlsx"
    databook = xlrd.open_workbook(path)
    datasheet = databook.sheet_by_index(0)
    
    
    
    
    
    
    time1 = ['0.5', '1', '2', '4', '8', '15', '30', '60', '120', '240', '480', '960', '1440']
    time = ['30', '60', '120', '240', '480', '900', '1800', '3600', '7200', '14400', '28800', '57600', '86400']
    calculateddata_hydrometer = [['Time(min)', 'Temperature(Celcius)', 'Hydrometer Reading', 'Density of Suspension(g/cc)', 'Density of Fluid(g/cc)', 'Height of Fall(cm)', 'Equivalent Diameter(mm)', 'Percent Finer than Equivalent Diameter(%)']]
    d = 0
    
    def densFluid(density, temperature, refT):
        dens = float(density[temperature.index(str(int(refT)))]) + ((float(refT)-float(temperature[temperature.index(str(int(refT)))]))/float(temperature[temperature.index(str(int(refT))) + 1]) - float(temperature[temperature.index(str(int(refT)))]))*(float(density[temperature.index(str(int(refT))) + 1]) - float(density[temperature.index(str(int(refT)))]))
        return dens
    def densSusp(hydrometerType, rm, meniscus):
        if hydrometerType == '151H':
            denssusp = (rm + meniscus) * 0.99821
        elif hydrometerType == '152H':
            denssusp = 0.99821 + (rm + meniscus) * 1.65/2650
        else:
            print('The Type of the Hydrometer is wrong!!')
        return denssusp
    def heightofFall(h2, h1, rm, meniscus, r2, r1, volumeHydrometer, areaCyl):
        heightoffall = h1 + ((h2 - h1) * (rm + meniscus - r1)) / (r2 - r1) - (volumeHydrometer / (2 * areaCyl))
        return heightoffall
    def diameter(h2, h1, rm, meniscus, r2, r1, volumeHydrometer, areaCyl, viscosity, temperature, density, refT, gs, t):
        dim = ((18*heightofFall(h2, h1, rm, meniscus, r2, r1, volumeHydrometer, areaCyl)*float(viscosity[temperature.index(str(int(refT)))]) + viscosityFluid(refT, temperature, viscosity)) / (((gs*0.99821) - densFluid(density, temperature, refT)) * 980.7 * float(t)))**0.5
        return dim
    def finerDiameter(gs, hydrometerType, rm, meniscus, density, temperature, totalDry, refT, samplePercent):
        finerdim = (((gs * 998.21 * (densSusp(hydrometerType, rm, meniscus) - densFluid(density, temperature, refT))) / (((gs * 0.99821) - densFluid(density, temperature, refT))*totalDry)))*samplePercent
        return finerdim
    def viscosityFluid(refT, temperature, viscosity):
        viscos = ((float(refT)-float(temperature[temperature.index(str(int(refT)))]))/float(temperature[temperature.index(str(int(refT))) + 1]) - float(temperature[temperature.index(str(int(refT)))]))*(float(viscosity[temperature.index(str(int(refT))) + 1]) - float(viscosity[temperature.index(str(int(refT)))]))
        return viscos
    
    #Taking data from excel
    totalDry = float(datasheet.cell_value(3, 13))
    gs = float(datasheet.cell_value(4, 13))
    maxGrainSize = float(datasheet.cell_value(5, 13))
    samplePercent = float(datasheet.cell_value(6, 13))
    refDens = float(datasheet.cell_value(7, 13))
    
    hydrometerType = str(datasheet.cell_value(10, 13))
    volumeHydrometer = float(datasheet.cell_value(11, 13))
    areaCyl = float(datasheet.cell_value(12, 13))
    meniscus = float(datasheet.cell_value(13, 13))
    r1 = float(datasheet.cell_value(14, 13))
    r2 = float(datasheet.cell_value(15, 13))
    h1 = float(datasheet.cell_value(16, 13))
    h2 = float(datasheet.cell_value(17, 13))
    
    for i in range(13):
        refT = float(datasheet.cell_value(i+21, 13))        
        rm = float(datasheet.cell_value(i+21, 14))
        calculateddata_hydrometer.append([])
        t = time[i]
        calculateddata_hydrometer[i+1].append(str(time1[i]))
        calculateddata_hydrometer[i+1].append(str(refT))
        calculateddata_hydrometer[i+1].append(str(rm))
        calculateddata_hydrometer[i+1].append(str(densSusp(hydrometerType, rm, meniscus)))
        calculateddata_hydrometer[i+1].append(str(densFluid(density, temperature, refT)))
        calculateddata_hydrometer[i+1].append(str(heightofFall(h2, h1, rm, meniscus, r2, r1, volumeHydrometer, areaCyl)))
        calculateddata_hydrometer[i+1].append(str(diameter(h2, h1, rm, meniscus, r2, r1, volumeHydrometer, areaCyl, viscosity, temperature, density, refT, gs, t)))
        calculateddata_hydrometer[i+1].append(str(finerDiameter(gs, hydrometerType, rm, meniscus, density, temperature, totalDry, refT, samplePercent)))
    
    x_hydr = []
    y_hydr = []
    for i in range(13):
        x_hydr.append(float(calculateddata_hydrometer[i+1][6]))
    for i in range(13):
        y_hydr.append(float(calculateddata_hydrometer[i+1][7]))
    x_hydr.reverse()
    y_hydr.reverse()
    
    #------------0.002 and 0.475-------------
    spline = CubicSpline(x_hydr,y_hydr,bc_type='natural')
    
    x_spline_hydr = np.linspace(x_hydr[0], x_hydr[-1], 39)
    y_spline_hydr = spline(x_spline_hydr)
    
    clay = spline(0.002)
    silt = spline(0.074)
    
    #------------EXCEL PART----------------
    excellinphydrometer1 = ['Total Dry Mass Tested (g)', 'Specific Gravity', 'Maximum Grain Size (mm)', 'Total Sample Percentage', 'Density of the Referenced Fluid (g/cc)']
    excellouthydrometer1 = [totalDry, gs, maxGrainSize, samplePercent, refDens]
    excellinphydrometer2 = ['Hydrometer Type', 'Volume of the Hydrometer (cm^3)', 'Area of the Cylinder (cm^2)', 'Meniscus Correction', 'r1', 'r2', 'H1', 'H2']
    excellouthydrometer2 = [volumeHydrometer, areaCyl, meniscus, r1, r2, h1, h2]
    
    excell = xlsxwriter.Workbook('Hydrometer Data ' + str(datetime.now().year) + '.' + str(datetime.now().month) + '.' + str(datetime.now().day) + ' ' + str(datetime.now().hour) + '.' + str(datetime.now().minute) + '.' + str(datetime.now().second) + '.xlsx')
    sheet = excell.add_worksheet('Hydrometer Data Results')
    
    bold_format = excell.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})
    
    normal_format = excell.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})
    
    sheet.merge_range('A1:H1', 'Hydrometer Test Results', bold_format)
    for i in range(8):
        sheet.write(1, i, calculateddata_hydrometer[0][i], bold_format)
    
    for i in range(13):
        for j in range(8):
            sheet.write_number(i+2, j, float("%.4f" % float(calculateddata_hydrometer[i+1][j])), normal_format)
    
    sheet.merge_range('A17:B17', 'Basic Properties', bold_format)
    sheet.merge_range('D17:E17', 'Hydrometer Properties', bold_format)
    
    sheet.write(17, 4, hydrometerType, normal_format)
    for i in range(5):
        sheet.write(i+17, 0, excellinphydrometer1[i], bold_format)
        sheet.write_number(i+17, 1, float("%.4f" % float(excellouthydrometer1[i])), normal_format)
    
    for i in range(8):
        sheet.write(i+17, 3, excellinphydrometer2[i], bold_format)
    for i in range(7):
        sheet.write_number(i+18, 4, float("%.4f" % float(excellouthydrometer2[i])), normal_format)
        
    excell.close()
    
    #-------------PLOTTING THE DISTRIBUTION TABLE--------------
    plt.plot(x_spline_hydr, y_spline_hydr, color = 'b')
    
    plt.axhline(linewidth=2, color='k')
    
    
    plt.title('Grain Size Distribution by Sedimentation')
    plt.xlabel('Particle Size')
    plt.ylabel('%finer')
    plt.xscale('log', basex = 10)
    plt.grid(True, which="both")
    plt.gca().spines['top'].set_visible(False)
    plt.gca().spines['right'].set_visible(False)
    plt.gca().set_xlim([0.0001, 100])
    plt.gca().set_ylim([0, 103])
    
    plt.axvspan(0, 0.002, facecolor = '#AAD1D7')
    plt.axvspan(0.002, 0.074, facecolor = '#A2A785')
    plt.axvspan(0.074, 4.76, facecolor = '#8FFF64')
    plt.axvspan(4.76, 76.2, facecolor = '#E5ED27')
                
    plt.text(0.0005, 90, 'CLAY', horizontalalignment='center')
    plt.text(0.014, 90, 'SILT', horizontalalignment='center')
    plt.text(0.417, 90, 'SAND', horizontalalignment='center')
    plt.text(20.48, 90, 'GRAVEL', horizontalalignment='center')
    
    plt.show()


"""-------------------ATTERBERG CONSISTENCY-------------------"""

if test_wanted == 4:
    #------------OPENING EXCEL--------------
    path = "SoilMechanicsDataSet.xlsx"
    databook = xlrd.open_workbook(path)
    datasheet = databook.sheet_by_index(0)
    
    
    
    
    


    calculateddata_atterberg = [['Type of Test', 'Liquid Limit', 'Plastic Limit'], 
                                ['Number of Drops'], 
                                ['Mass of Container + Wet Soil (g)'], 
                                ['Mass of Container + Dry Soil (g)'], 
                                ['Mass of Container (g)'], 
                                ['Mass of Moisture (g)'], 
                                ['Mass of Dry Soil (g)'], 
                                ['Moisture Content (%)'], 
                                ['Limit (%)']]
    
    for i in range(4):
        for j in range(5):
            datas = str(datasheet.cell_value(i+3, j+18))
            calculateddata_atterberg[i+1].append(datas)
    
    for i in range(5):
        massmoist = float(calculateddata_atterberg[2][i+1]) - float(calculateddata_atterberg[3][i+1])
        calculateddata_atterberg[5].append(str(massmoist))
        massdry = float(calculateddata_atterberg[3][i+1]) - float(calculateddata_atterberg[4][i+1])
        calculateddata_atterberg[6].append(str(massdry))
        contmoist = (massmoist *100) / massdry
        calculateddata_atterberg[7].append(str(contmoist))
    
    plaslimit = (float(calculateddata_atterberg[7][4]) + float(calculateddata_atterberg[7][5])) / 2
    
    
    x_attr = []
    y_attr = []
    
    if float(calculateddata_atterberg[7][1]) < float(calculateddata_atterberg[7][2]):
        if float(calculateddata_atterberg[7][1]) < float(calculateddata_atterberg[7][3]):
            x_attr.append(float(calculateddata_atterberg[7][1]))
            y_attr.append(float(calculateddata_atterberg[1][1]))
            if float(calculateddata_atterberg[7][2]) < float(calculateddata_atterberg[7][3]):
                x_attr.append(float(calculateddata_atterberg[7][2]))
                y_attr.append(float(calculateddata_atterberg[1][2]))
                x_attr.append(float(calculateddata_atterberg[7][3]))
                y_attr.append(float(calculateddata_atterberg[1][3]))
            else:
                x_attr.append(float(calculateddata_atterberg[7][3]))
                y_attr.append(float(calculateddata_atterberg[1][3]))
                x_attr.append(float(calculateddata_atterberg[7][2]))
                y_attr.append(float(calculateddata_atterberg[1][2]))
        else:
            x_attr.append(float(calculateddata_atterberg[7][3]))
            y_attr.append(float(calculateddata_atterberg[1][3]))
            x_attr.append(float(calculateddata_atterberg[7][1]))
            y_attr.append(float(calculateddata_atterberg[1][1]))
            x_attr.append(float(calculateddata_atterberg[7][2]))
            y_attr.append(float(calculateddata_atterberg[1][2]))
    
    elif float(calculateddata_atterberg[7][2]) < float(calculateddata_atterberg[7][3]):
        x_attr.append(float(calculateddata_atterberg[7][2]))
        if float(calculateddata_atterberg[7][1]) < float(calculateddata_atterberg[7][3]):
            x_attr.append(float(calculateddata_atterberg[7][1]))
            y_attr.append(float(calculateddata_atterberg[1][1]))
            x_attr.append(float(calculateddata_atterberg[7][3]))
            y_attr.append(float(calculateddata_atterberg[1][3]))
        else:
            x_attr.append(float(calculateddata_atterberg[7][3]))
            y_attr.append(float(calculateddata_atterberg[1][3]))
            x_attr.append(float(calculateddata_atterberg[7][1]))
            y_attr.append(float(calculateddata_atterberg[1][1]))
    else:
        x_attr.append(float(calculateddata_atterberg[7][3]))
        if float(calculateddata_atterberg[7][1]) < float(calculateddata_atterberg[7][2]):
            x_attr.append(float(calculateddata_atterberg[7][1]))
            y_attr.append(float(calculateddata_atterberg[1][1]))
            x_attr.append(float(calculateddata_atterberg[7][2]))
            y_attr.append(float(calculateddata_atterberg[1][2]))
        else:
            x_attr.append(float(calculateddata_atterberg[7][2]))
            y_attr.append(float(calculateddata_atterberg[1][2]))
            x_attr.append(float(calculateddata_atterberg[7][1]))
            y_attr.append(float(calculateddata_atterberg[1][1]))
    
    
    plt.scatter(y_attr, x_attr)
    
    # calc the trendline
    z = np.polyfit(y_attr, x_attr, 1)
    p = np.poly1d(z)
    plt.plot(y_attr,p(y_attr),"r--")
    # the line equation:
    liqlimit = z[0] * 25 + z[1]
    
    #--------------EXCEL PART---------------
    attrexcell = xlsxwriter.Workbook('Atterberg Data ' + str(datetime.now().year) + '.' + str(datetime.now().month) + '.' + str(datetime.now().day) + ' ' + str(datetime.now().hour) + '.' + str(datetime.now().minute) + '.' + str(datetime.now().second) + '.xlsx')
    attrsheet = attrexcell.add_worksheet('Atterberg Data Results')
    
    bold_format = attrexcell.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})
    
    normal_format = attrexcell.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})
    
    attrsheet.merge_range('A1:F1', 'Atterberg (Consistency) Limits Test Results', bold_format)
    attrsheet.merge_range('B3:D3', 'Liquid Limit', bold_format)
    attrsheet.merge_range('E3:F3', 'Plastic Limit', bold_format)
    
    for i in range(9):
        attrsheet.write(i+2, 0, calculateddata_atterberg[i][0], bold_format)
    for i in range(2):
        attrsheet.write(3, i+4, calculateddata_atterberg[1][i+4], bold_format)
    
    for i in range(6):
        for j in range(5):
            attrsheet.write_number(i+4, j+1, float("%.4f" % float(calculateddata_atterberg[i+2][j+1])), normal_format)
    for i in range(3):
        attrsheet.write_number(3, i+1, float("%.4f" % float(calculateddata_atterberg[1][i+1])), normal_format)
    
    attrsheet.merge_range('B11:D11', str(int(liqlimit))+' %', bold_format)
    attrsheet.merge_range('E11:F11', str(int(plaslimit))+ ' %', bold_format)
    
    attrexcell.close()



    plt.title('Atterberg Liquid Limit Analysis')
    plt.xlabel('Number of Drops')
    plt.ylabel('Moisture Content (%)')
    plt.xscale('log', basex = 10)
    plt.grid(True, which="both")
    plt.gca().set_xlim([1, 100])
    plt.gca().set_ylim([0, 100])

    plt.show()




"""-------------------STANDARD PROCTOR-------------------"""
if test_wanted == 5:
    #------------OPENING EXCEL--------------
    path = "SoilMechanicsDataSet.xlsx"
    databook = xlrd.open_workbook(path)
    datasheet = databook.sheet_by_index(0)
    
    
    
    totalmass = float(datasheet.cell_value(2, 26))
    gs_sample = float(datasheet.cell_value(3, 26))
    coarser95 = float(datasheet.cell_value(4, 26))
    wcoarse = float(datasheet.cell_value(5, 26))
    gs_coarse = float(datasheet.cell_value(6, 26))
    height_mold = float(datasheet.cell_value(2, 29))
    diam_mold = float(datasheet.cell_value(3, 29))
    vol_mold = 3.14159265 * height_mold * (diam_mold/2)**2
    
    calculateddata_proctor1 = [['Total Mass of Sample (g)', 'Specific Gravity of the Sample', 'Coarser than 9.5 mm (%)', 'Water Content of Coarse (%)', 'Specific Gravity of the Removed Fraction', 'Height of the Mold (cm)', 'Diameter of Mold (cm)', 'Volume of the Mold (cm^3)'], 
                                [totalmass, gs_sample, coarser95, wcoarse, gs_coarse, height_mold, diam_mold, vol_mold]]
    
    calculateddata_proctor2 = [['Test Number', '1', '2', '3', '4', '5'], 
                               ['Mass of Mold + Base (g)'], 
                               ['Mass of Mold + Base + Compacted Soil (g)'], 
                               ['Mass of Compacted Soil (g)'], 
                               ['Bulk Density (g/cc)'], 
                               ['Dry Density (g/cc)'], 
                               ['Corrected Dry Density (g/cc)']]
    
    calculateddata_proctor3 = [['Test Number', '1', '2', '3', '4', '5'], 
                               ['Container + Wet Sample (g)'], 
                               ['Container + Dry Sample (g)'], 
                               ['Mass of Container (g)'], 
                               ['Mass of Moisture (g)'], 
                               ['Dry Mass (g)'], 
                               ["Moisture Content (%)"],
                               ["Corrected Moisture Content (%)"]]
    
    for i in range(5):
        Mmb = float(datasheet.cell_value(9, i+26))
        Mmbc = float(datasheet.cell_value(10, i+26))
        calculateddata_proctor2[1].append(str(Mmb))
        calculateddata_proctor2[2].append(str(Mmbc))
    
    for i in range(5):
        cpluswet = float(datasheet.cell_value(14, i+26))
        cplusdry = float(datasheet.cell_value(15, i+26))
        onlyc = float(datasheet.cell_value(16, i+26))
        calculateddata_proctor3[1].append(str(cpluswet))
        calculateddata_proctor3[2].append(str(cplusdry))
        calculateddata_proctor3[3].append(str(onlyc))
    
    for i in range(5):
        calculateddata_proctor3[4].append(str(float(calculateddata_proctor3[1][i+1]) - float(calculateddata_proctor3[2][i+1])))
        calculateddata_proctor3[5].append(str(float(calculateddata_proctor3[2][i+1]) - float(calculateddata_proctor3[3][i+1])))
        calculateddata_proctor3[6].append(str((float(calculateddata_proctor3[4][i+1]) / float(calculateddata_proctor3[5][i+1]))*100))
        calculateddata_proctor3[7].append(str((((wcoarse/100)*(coarser95/100))+((1-(coarser95/100)) * (float(calculateddata_proctor3[6][i+1])/100)))*100))
    
    for i in range(5):
        calculateddata_proctor2[3].append(str(float(calculateddata_proctor2[2][i+1]) - float(calculateddata_proctor2[1][i+1])))
        calculateddata_proctor2[4].append(str(float(calculateddata_proctor2[3][i+1])/vol_mold))
        calculateddata_proctor2[5].append(str(((float(calculateddata_proctor2[4][i+1])*100)/(100+(float(calculateddata_proctor3[6][i+1]))))))
        calculateddata_proctor2[6].append(str((float(calculateddata_proctor2[5][i+1])*gs_coarse*1)/((float(calculateddata_proctor2[5][i+1])*(coarser95/100))+(gs_coarse*1*(1-(coarser95/100))))))
    
    
    x_proctor = []
    y_proctor = []
    for i in range(5):
        x_proctor.append(float(calculateddata_proctor3[7][i+1]))
        y_proctor.append(float(calculateddata_proctor2[6][i+1]))
    
    spline = CubicSpline(x_proctor,y_proctor,bc_type='natural')
    
    x_spline_proctor = np.linspace(x_proctor[0], x_proctor[-1], 1000)
    y_spline_proctor = spline(x_spline_proctor)
    
    srcurves = np.linspace(0, 40, 1000)
    
    sr100list = []
    sr90list = []
    sr80list = []
    xsrlist = []
    
    for i in range(5):
        sr100 = gs_sample/(1+(float(calculateddata_proctor3[7][i+1])*gs_sample/100))
        sr100list.append(sr100)
        sr90 = gs_sample/(1+(float(calculateddata_proctor3[7][i+1])*gs_sample/90))
        sr90list.append(sr90)
        sr80 = gs_sample/(1+(float(calculateddata_proctor3[7][i+1])*gs_sample/80))
        sr80list.append(sr80)
        xsrlist.append(float(calculateddata_proctor3[7][i+1]))
    
    
    #-----------EXCEL PART----------
    proctorexcell = xlsxwriter.Workbook('Standard Proctor Compaction Data ' + str(datetime.now().year) + '.' + str(datetime.now().month) + '.' + str(datetime.now().day) + ' ' + str(datetime.now().hour) + '.' + str(datetime.now().minute) + '.' + str(datetime.now().second) + '.xlsx')
    proctorsheet = proctorexcell.add_worksheet('Standard Proctor Data Results')
    
    bold_format = proctorexcell.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})
    
    normal_format = proctorexcell.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})
    
    proctorsheet.merge_range('A1:F1', 'Standard Proctor Compaction Test Results', bold_format)
    proctorsheet.merge_range('A17:F17', 'Moisture Content', bold_format)
    
    for i in range(5):
        proctorsheet.write(i+2, 0, calculateddata_proctor1[0][i], bold_format)
        proctorsheet.write_number(i+2, 1, float("%.4f" % float(calculateddata_proctor1[1][i])), normal_format)
    for i in range(2):
        proctorsheet.write(i+2, 3, calculateddata_proctor1[0][i+5], bold_format)
        proctorsheet.write_number(i+2, 4, float("%.4f" % float(calculateddata_proctor1[1][i+5])), normal_format)
    for i in range(7):
        proctorsheet.write(i+8, 0, calculateddata_proctor2[i][0], bold_format)
    for i in range(5):
        proctorsheet.write_number(8, i+1, float("%.4f" % float(calculateddata_proctor2[0][i+1])), bold_format)
    for i in range(6):
        for j in range(5):
            proctorsheet.write_number(i+9, j+1, float("%.4f" % float(calculateddata_proctor2[i+1][j+1])), normal_format)
    for i in range(8):
        proctorsheet.write(i+17, 0, calculateddata_proctor3[i][0], bold_format)
    for i in range(5):
        proctorsheet.write_number(17, i+1, float("%.4f" % float(calculateddata_proctor3[0][i+1])), bold_format)
    for i in range(7):
        for j in range(5):
            proctorsheet.write_number(i+18, j+1, float("%.4f" % float(calculateddata_proctor3[i+1][j+1])), normal_format)
            
    proctorexcell.close()
    
    
    #-------------PLOTING-------------
    plt.plot(x_spline_proctor, y_spline_proctor, color = 'b',  label = 'Proctor Test Result')
    plt.plot(xsrlist, sr100list, color = 'y', linestyle='--', label = 'Sr 100')
    plt.plot(xsrlist, sr90list, color = 'r', linestyle='--', label = 'Sr 90')
    plt.plot(xsrlist, sr80list, color = 'k', linestyle='--', label = 'Sr 80')
    
    
    plt.title('Standard Proctor Test Analysis')
    plt.xlabel('Corrected Water Content (%)')
    plt.ylabel('Corrected Density (g/cc)')
    plt.grid(True, which="both")
    plt.gca().spines['top'].set_visible(False)
    plt.gca().spines['right'].set_visible(False)
    
    plt.legend(loc = 'upper right')
    
    plt.gca().set_xlim([10, 25])
    plt.gca().set_ylim([1, 2.5])
    
    plt.show()






