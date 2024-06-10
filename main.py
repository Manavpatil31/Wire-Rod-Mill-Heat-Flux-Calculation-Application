import linecache                      #Module to fetch a particular line from the text file
import pandas as pd                   #Module to fetch excel sheet and creating DataFrames
import matplotlib.pyplot as plt       #Module to plot graphs
# import tkinter as tk                  #Module used to work on GUI(Graphical User Interface)
from tkinter import messagebox        #To display message-box using tkinter


#********************* Setting icon and title for the application ***********
                                #(Future Use)
# app = tk.Tk()
# app.title('App Title')
# app.iconbitmap('icon.ico')

#*********************


#********************* Fetching Manual Entry from text **********************

file = open('2 Manual Entry.txt', 'a')

WD1 = linecache.getline('2 Manual Entry.txt', 8)       #WD = Wire Diameter(mm)
WD = float(WD1)
if(WD>=5 and WD<=15):
    WD = WD
else:
    messagebox.showerror('Error', 'Enter the correct Wire Diameter, it ranges from 5-15mm.')

ST1 = linecache.getline('2 Manual Entry.txt', 9)       #ST = Wire Rod Surface Temperature(C)
ST = int(ST1)
if(ST>=650 and ST<=1000):
    ST = ST
else:
    messagebox.showerror('Error', 'Enter the correct Wire Rod Surface Temperature, it ranges from 650-1000C.')

AAT1 = linecache.getline('2 Manual Entry.txt', 10)     #AAT = Air Ambient Temperature(C)
AAT = int(AAT1)
if(AAT>=20 and AAT<=50):
    AAT = AAT
else:
    messagebox.showerror('Error', 'Enter the correct Air Ambient Temperature, it ranges from 20-50C.')

TE1 = linecache.getline('2 Manual Entry.txt', 11)      #TE = Total Emissivity including shape factor
TE = float(TE1)
if(TE>=0 and TE<=1):
    TE = TE
else:
    messagebox.showerror('Error','Enter the correct value for Total Emissivity, it ranges from 0-1.')

AAV1 = linecache.getline('2 Manual Entry.txt', 12)      #AAV = Avearge Air Velocities(m/s)
AAV2 = AAV1.split(',')
AAV3 = list(AAV2)
AAV = [eval(i) for i in AAV3]
SBC = 0.0000000556                                   #SBC = Stefan-Boltzmann Constant
AG = 9.806                                           #AG = Acceleration due to Gravity

#**************************************************************************************

#*********** Calculating Film temperature *************
T_film = (ST+AAT)/2                                  #T_film = Film Temperature
T_film1 = round(T_film, 2)

#*************** Fetching Closest temperature from Air Property Sheet in Excel**************************

Air_Property = pd.read_excel('1 Air Property.xlsx', sheet_name = 'Air Property')
Air_Property_closest_T = Air_Property.iloc[(Air_Property['Temp']-T_film1).abs().argsort()[0:2]]
dt = Air_Property_closest_T['Temp'].tolist()
dt_1 = dt[0]
dt_2 = dt[1]

#Calculating Thermal Conductivity of Air
dtc = Air_Property_closest_T['Thermal Conductivity'].tolist()
dtc_1 = dtc[0]                              #Lower value of Thermal Conductivity
dtc_2 = dtc[1]                              #Higher value of Thermal Conductivity
TCA = ((dtc_2-dtc_1)/(dt_2-dt_1)*(T_film1-dt_2)+dtc_2)
TCA1 = round(TCA, 4)

#Calculating Kinematic Viscosity
dkv = Air_Property_closest_T['Kinematic Viscosity'].tolist()
dkv_1 = dkv[0]
dkv_2 = dkv[1]
KV = ((dkv_2-dkv_1)/(dt_2-dt_1)*(T_film1-dt_2)+dkv_2)
KV1 = round(KV, 4)

#Calculating Thermal Diffusivity
dtd = Air_Property_closest_T['Thermal Diffusivity'].tolist()
dtd_1 = dtd[0]
dtd_2 = dtd[1]
TD = ((dtd_2-dtd_1)/(dt_2-dt_1)*(T_film1-dt_2)+dtd_2)
TD1 = round(TD, 4)

#Calculating Rayleigh Number
RN = AG*(ST-AAT)*((WD*10**-3))**3/(KV1*TD1*(T_film1+273))
RN1 = round(RN, 4)

#************************** Calculating For C and N ************************************

C = float()
N = float()
if(RN1>0.1 and RN1<100):
    C = 1.02
    N = 0.148
elif(RN1>101 and RN1<10000):
    C = 0.85
    N = 0.188
elif(RN>10001 and RN1<10000000):
    C = 0.48
    N = 0.25
else:
    print('Not Valid')

#*************************************************************************************

#********** Calculating Heat Transfer Coefficient-1 **********

#Calculating Force Convection Heat Transfer Coefficient-1
HFC = [(0.683*(TCA1/(WD*10**-3))*(i*(WD*10**-3)/KV1)**0.466*(KV1/TD1)**0.33) for i in AAV]
HFC = [round(item, 4) for item in HFC]

#Calculating Radiation Heat Transfer Coefficient-1
HR = TE*SBC*(((ST+273)**4-(AAT+273)**4)/((ST+273)-(AAT+273)))
HR = round(HR, 4)
HR_L = [HR] * len(HFC)

#Calculating Natural Convection Heat Transfer Coefficient-1
HNC = TCA1/(WD*(10**-3))*C*(RN1)**N
HNC = round(HNC, 4)
HNC_l = [HNC] * len(HFC)


#******************** Calculating Heat Transfer Coefficient-2 *****************

#Calculating Force Convection Heat Transfer Coefficient-2
HFC_2 = [(j*(T_film1+273)/(10**3)) for j in HFC]
HFC_2 = [round(item2, 4) for item2 in HFC_2]

#Calculating Radiation Heat Transfer Coefficient-2
HR_2 = HR*(T_film1+273)/(10**3)
HR_2 = round(HR_2, 4)
HR_2l = [HR_2] * len(HFC_2)

#Calculating Natural Convection Heat Transfer Coefficient-2
HNC_2 = HNC*(T_film1+273)/(10**3)
HNC_2 = round(HNC_2, 4)
HNC_2l = [HNC_2] * len(HFC_2)


#**************************** Plotting the Graph **********************************

#Plotting for Force Convection Heat Transfer Coefficient
x1 = AAV
y1 = HFC_2
plt.plot(x1,y1, marker = 'o', label = 'Force Convection HTC')

#Plotting for Radiation Heat Transfer Coefficient
x2 = AAV
y2 = HR_2l
plt.plot(x2,y2, marker = 'o', label = 'Radiation HTC')

#Plotting for Natural Convection Heat Transfer Coefficient
x3 = AAV
y3 = HNC_2l
plt.plot(x3,y3,marker = 'o', label = 'Natural Convection HTC')


plt.xlabel('Average Air Velocity')
plt.ylabel('Heat Transfer Coefficient')
plt.title('Heat Transfer Model')
plt.legend()
done = plt.savefig('5 Heat Transfer Graph.png')

#*************************************************************************************

#****************** Generating an excel sheet of the data calculated *********************

Data = pd.DataFrame({'Wire Diameter(mm)': WD,'WR Surface Temp(C)': ST,'Air Ambient Temp(C)': AAT,'Average Air Velocity(m/s)': AAV,'Total Emissivity':TE,'Force Convection HTC':HFC_2,'Radiation HTC':HR_2l,'Natural Convection HTC':HNC_2l})
file_name = '4 Heat Transfer Result.xlsx'
writer = pd.ExcelWriter(file_name, engine='xlsxwriter', mode = 'w')
Data.to_excel(writer, sheet_name = 'Result')
workbook = writer.book
worksheet = writer.sheets['Result']
worksheet.insert_image('K5', '5 Heat Transfer Graph.png')
writer.close()

#*********************** Displaying Message-box to check the completion of the Process ************
messagebox.showinfo("Confirmation", "Process Completed!")
