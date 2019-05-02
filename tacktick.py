def main():
    import pandas as pd
    import xlrd
    import math
    import time
    import datetime
    from datetime import timedelta


    data = pd.read_excel(r'D:\iaind\Documents\Sailing\testData1.xlsx')
    [(ix, k, v) for ix, row in data.iterrows() for k, v in row.items()]
    print (data)



    # df1 = pd.read_excel (r'D:\iaind\Documents\Sailing\testData1.xlsx',"Sheet1") 
    # df1 = df1.dropna()
    # data 
    # heading = df1["heading"].values.tolist()

    # df2 = pd.read_excel (r'D:\iaind\Documents\Sailing\tacktickTestData.xlsx',"Sheet2") 
    # df2 = df2.dropna() 
    # newHeading = df2["heading"].values.tolist()

    # print(heading)
    # print(newHeading)
    # a = len(newHeading)
    # for i in range(a):
    #     heading.append(newHeading.pop(0))
    #     if (len(heading)>=20):
    #         heading.pop(0)
    #     average = sum(heading)/len(heading)
    #     print("Average:")
    #     print(round(average,2))
    #     print("Shift:")
        
    #     print(round(heading[-1] - average,2))

    #data point tuples
    pointA = (0, 0,25.90)
    pointB = (0, -1, 45.92)

    datetimeFormat = '%Y-%m-%dT%H:%M:%S.%fZ'
    d2 = '2019-05-01T20:51:45.924Z'
    d1 = '2019-05-01T20:51:25.904Z'
    diff = datetime.datetime.strptime(d2, datetimeFormat) - datetime.datetime.strptime(d1, datetimeFormat)

    print (diff)
    


    if (type(pointA) != tuple) or (type(pointB) != tuple):
        raise TypeError("Only tuples are supported as arguments")

    #Lat and Long 
    lat1 = math.radians(pointA[0])
    lat2 = math.radians(pointB[0])
    long1 = math.radians(pointA[1])
    long2 = math.radians(pointB[1])
    
    diffLat = lat2 - lat1
    diffLong = long2 - long1
    
    #time different from points 
    time = pointB[2] - pointA[2] #seconds
    time = time / 3600 #hours

    #Calcuations for distance travlled
    R = 6373.0
    a = math.sin(diffLat / 2)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(diffLong / 2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

    #distance travelled in KM
    distance = R * c
    #Convert distance to Knots
    distanceKnots  = distance * 0.539957
    #Calculate spped in Knots
    speed = distanceKnots / time


    

    x = math.sin(diffLong) * math.cos(lat2)
    y = math.cos(lat1) * math.sin(lat2) - (math.sin(lat1) * math.cos(lat2) * math.cos(diffLong))

    initial_bearing = math.atan2(x, y)
 

    # Now we have the initial bearing but math.atan2 return values
    # from -180° to + 180° which is not what we want for a compass bearing
    # The solution is to normalize the initial bearing as shown below
    
    initial_bearing = math.degrees(initial_bearing)
    compass_bearing = (initial_bearing + 360) % 360

    print ("bearing: ", round(compass_bearing,2))
    print("Distance:", round(distance,2), "km")
    print("Speed: ", round(speed,2), "Knots")



if __name__ == "__main__":
    main()

