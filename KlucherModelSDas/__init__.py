import xlsxwriter
import math

constants = [
    [1202, 0.141, 0.103], [1187, 0.142, 0.104], [1164, 0.149, 0.109], [1130, 0.164, 0.120], [1106, 0.177, 0.130], [1092, 0.185, 0.137],
    [1093, 0.186, 0.138], [1107, 0.182, 0.134], [1136, 0.165, 0.121], [1136, 0.152, 0.111], [1190, 0.144, 0.106], [1204, 0.141, 0.103]
]

def checkYear(year):
    if (year % 4) == 0:  
        if (year % 100) == 0:  
            if (year % 400) == 0:
                return True 
            else:      
                return False
        else:  
            return True  
    else:  
        return False

def klucher(year, lat, tilt, azi, interval):
    n = 0        #day of the year
    hra = 0      #hour angle
    dec = 0      #declination angle
    tilt     #tilt angle
    azi     #azimuth angle
    cosQ = 0     #angle of incidence of solar radiation
    cosQz = 0    #zenith angle
    F = 0        #correction coefficient
        
    Ig = 0       #Hourly Global Radiation
    Ib = 0       #Hourly Beam Radiation
    Id = 0       #Hourly Diffused Radiation
    It = 0       #Solar Radiation on Tilted Surface

    # lat = float(input("Enter Latitude: "))      #latitude     
    # year = int(input("Enter a year: ")) 
    # tilt = float(input("Enter tilt: "))      #tilt angle     
    # azi = float(input("Enter azi: "))      #azimuth angle
    # interval = int(input("Enter the interval: "))    #interval of mins for data representation

    workbook = xlsxwriter.Workbook('data.xlsx')
    worksheet = workbook.add_worksheet()
    row = 0
    column = 0

    worksheet.write(row, column, 'Year')
    column += 1 
    worksheet.write(row, column, 'Latitude')
    column += 1
    worksheet.write(row, column, 'Tilt Angle')
    column += 1
    worksheet.write(row, column, 'Azimuth Angle')
    row +=1
    column=0

    worksheet.write(row, column, year)
    column += 1 
    worksheet.write(row, column, lat)
    column += 1
    worksheet.write(row, column, tilt)
    column += 1
    worksheet.write(row, column, azi)
    row +=1
    column=0

    #Converting degrees into radian
    lat *=3.14159/180
    dec *=3.14159/180
    tilt *=3.14159/180
    azi *=3.14159/180

    worksheet.write(row, column, 'n')
    column += 1 
    worksheet.write(row, column, 'Time')
    column += 1
    worksheet.write(row, column, 'Hour Angle')
    column += 1
    worksheet.write(row, column, 'Declination')
    column += 1
    worksheet.write(row, column, 'cosQ')
    column += 1 
    worksheet.write(row, column, 'cosQz')
    column += 1
    worksheet.write(row, column, 'Ib')
    column += 1
    worksheet.write(row, column, 'Id')
    column += 1
    worksheet.write(row, column, 'Ig')
    column += 1
    worksheet.write(row, column, 'F')
    column += 1
    worksheet.write(row, column, 'It')
    row +=1
    column=0

    val = [0,31,28,31,30,31,30,31,31,30,31,30,31]

    for month in range(1, 13):
        ctr = 0

        if (checkYear(year) and month==2):
            ctr = val[month] +1
        else:
            ctr = val[month]

        for date in range(1,(ctr+1)):
            n = date + (month>1)*31 + (month>2)*28 + (month>3)*31 + (month>4)*30 + (month>5)*31 + (month>6)*30 + (month>7)*31 + (month>8)*31 + (month>9)*30 + (month>10)*31 + (month>11)*30

            n = n + (1*(checkYear(year) and month>2))     #if the year is leap year

            deg = (360*(284+n))/365
            deg*=(3.14159/180)
            dec = 23.45 * math.sin(deg)
            dec *= 3.14159/180
            count = 1
            for hour in range(5, 18):
                for min in range(0, 60, interval):
                    hra = 15 * (12 - (hour + min/60))
                    hra *=3.14159/18028

                    cosQ = math.sin(lat)*(math.sin(dec)*math.cos(tilt) + math.cos(dec)*math.cos(azi)*math.cos(hra)*math.sin(tilt)) + math.cos(lat)*(math.cos(dec)*math.cos(hra)*math.cos(tilt) - math.sin(dec)*math.cos(azi)*math.sin(tilt)) + math.cos(dec)*math.sin(azi)*math.sin(hra)*math.sin(tilt)

                    cosQz = math.sin(lat)*math.sin(dec) + math.cos(lat)*math.cos(dec)*math.cos(hra)
                    
                    Ibn = constants[(month-1)][0] *  math.exp(-constants[(month-1)][1]/cosQz)
                    Ib = Ibn * cosQz

                    Id = constants[(month-1)][2] * Ibn

                    Ig = Ib + Id
                    hra = hra/(3.14159/180)

                    F = 1 - ((Id/Ig) * (Id/Ig))

                    temp = tilt/(3.14159/180)      #temp is tilt in degrees
                    temp = temp/2
                    temp *= 3.14159/180 

                    sinQz = pow((1 - (cosQz*cosQz)), 0.5)

                    It = Ib * (cosQ/cosQz) + Id * ((1+math.cos(temp))/2) * (1 + (F * pow((math.sin(temp)),3))) * (1 + F * cosQz * cosQz *  sinQz * sinQz * sinQz)

                    hra = hra/(3.14159/180)
                    
                    time = str(hour)+' : '+str(min)

                    worksheet.write(row, column, n)
                    column += 1 
                    worksheet.write(row, column, time)
                    column += 1
                    worksheet.write(row, column, hra)
                    column += 1
                    worksheet.write(row, column, dec)
                    column += 1
                    worksheet.write(row, column, cosQ)
                    column += 1 
                    worksheet.write(row, column, cosQz)
                    column += 1
                    worksheet.write(row, column, Ib)
                    column += 1
                    worksheet.write(row, column, Id)
                    column += 1
                    worksheet.write(row, column, Ig)
                    column += 1
                    worksheet.write(row, column, F)
                    column += 1
                    worksheet.write(row, column, It)
                    row +=1
                    column=0

    workbook.close()

