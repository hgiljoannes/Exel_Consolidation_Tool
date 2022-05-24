# ExcelMonitoringFiles
This is a program to monitoring files in a specific folder that you select and check files in Excel, if the program find a excel file, it will copy the sheets into a new file root call ExcelRoot.xlsx.


When you execute the .exe program it will show you this windows

![image](https://user-images.githubusercontent.com/106103512/169869178-dd99e4bd-014a-4022-b222-cf3ae6a68345.png)

Then you have to select the folder that you want to check files

![image](https://user-images.githubusercontent.com/106103512/169871133-85bec6e8-1510-445a-87dc-7a67725aa922.png)

it will show you a windows, then select your folder to check files, can be any folder, for this ex. Excel
![image](https://user-images.githubusercontent.com/106103512/169873736-9ca92da4-1797-4dad-af22-87f24456ab18.png)

When you have selected your folder, it will start monitoring files
![image](https://user-images.githubusercontent.com/106103512/169875925-d2f3a3f9-f0e8-4bfe-a975-0f61e2f23cbd.png)

Then the program will create 3 files
![image](https://user-images.githubusercontent.com/106103512/169876304-49ff7245-e5d0-466c-9e7e-5c1373b3bafc.png)
it will create the 2 folders
- Processed: it will contains the processed excel files that have copy succesfully to the root excel file
- Not Aplicable: it will contains the files that are not excel files, any files except Excel files

ExcelRoot File: it will contains the copy sheets of the other excel files that have processed

#Tests
to make the test, just copy a excel file that contains data in the sheets

in this example we will copy this file to the Excel folder that we have selected before.
![image](https://user-images.githubusercontent.com/106103512/169879107-0d61357b-088b-4cf4-a27d-b0dee5a35256.png)
![image](https://user-images.githubusercontent.com/106103512/169879160-e896cdc8-4b22-465c-a7b0-e2d3d4b9db7a.png)

Ad you can see the Root Excel File not contains any data.
![image](https://user-images.githubusercontent.com/106103512/169879291-c1143bed-0ecb-4d87-a661-10b7cbd4f22f.png)

Then just copy the file, and the windows it will show you if the file processed or not.
![image](https://user-images.githubusercontent.com/106103512/169879514-332be371-880d-447f-b0f2-1fa57c940b45.png)
Result
![image](https://user-images.githubusercontent.com/106103512/169879550-848518eb-98a7-4fbf-95dd-38542f1d75c6.png)

And when you check the ExcelRoot.xlsx, it will copy the sheets of the Test File to ExcelRoot
![image](https://user-images.githubusercontent.com/106103512/169879824-ba0d6919-fd2b-4348-bae9-f685caba306f.png)
![image](https://user-images.githubusercontent.com/106103512/169879868-7527d223-eb3c-4d11-adc3-9c98366972a1.png)

And the file that have been processed, it will move to the processed folder
![image](https://user-images.githubusercontent.com/106103512/169879981-16bda4bf-e7f6-451f-a204-f3874a802b4e.png)

#Not Aplicable Files
If the file that we have created in the folder that we have monitoring is not a ExcelFile, it will show that is not aplicable and it will move it to the Not Aplicable Folder. (Any File that is not .xlsx

![image](https://user-images.githubusercontent.com/106103512/169946637-f438b60d-c0f9-45d7-919d-6afe45c8f609.png)






