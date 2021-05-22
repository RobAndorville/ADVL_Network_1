# ADVL_Network_1
### Andorville™ Network Software.



- - -
Copyright 2018 Signalworks Pty Ltd, ABN 26 066 681 598

Licensed under the Apache License, Version 2.0 (the "License");  
you may not use this file except in compliance with the License.  
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software  
distributed under the License is distributed on an "AS IS" BASIS,  
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.  
See the License for the specific language governing permissions and  
limitations under the License.



- - -


The Andorville™ Network software is used by Andorville™ application cells to exchange information. The Network shows a list of available application cells and projects and a list of projects that are currently connected.

#### Installation Notes

The setup files are contained in the zip file: Setup_ADVL_Network_1_0_1_22.Zip  
Press **Releases** on the right edge of this page to open the Releases page.  
The setup zip file is listed under **Assets**.

Unzip the directory: ADVL_Network_1_0_1_22  
The setup files are contained in this directory.  

Before running setup, the following URL must be reserved for the Message Service that is started when the Network software is run:  http://+:8733/ADVLService/

After the URL has been reserved, run the setup.exe file to install the Andorville™ Network software.

#### Reserve the URL for the Message Service on a Windows 10 Computer
To get access to the Message Service URL, right-click the start button and select Windows PowerShell (Admin).  
Enter the command:  
netsh http add urlacl url=http://+:8734/ADVLService/ user=Everyone  
If successful, this message will be shown: URL reservation successfully added

#### Delete a URL Reservation
To delete a URL reservation enter the command in the PowerShell:  
netsh http delete urlacl url=http://+:8733/ADVLService  
If successful, this message will be shown: URL reservation successfully deleted

#### Display a List of Active Connections
To display a list of active connections on a computer, enter the command in the PowerShell:  
netstat -o -n -a




