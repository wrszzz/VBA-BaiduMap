# VBA Getdistance with Baidu API
Useful VBA tools for logistics companies.
Manual:
1. Download and unzip VBA-JSON 2.3.1, import JsonConverter.bas to the module you created.
2. Import all other files to the module
3. Replace "Your API Key" with your api key. You can apply one at this website:https://lbsyun.baidu.com/
4. Open Reference...in the tool menu, and click Microsoft Scripting Runtime
5. Goback to the spread sheet and use any cell to record the origin and destination, use "=GetLatLng(loc1,loc2)" to get the longitude and latitude of the location, use "=getdistance(latlng1,latlng2)" to get the distance, use "=gettruckdistance(latlng1,latlng2)"to get the truck milage(premier API required).
