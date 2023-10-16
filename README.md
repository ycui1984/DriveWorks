# DriveWorks

Managing Google drive files could be time consuming and error prone. 
Imagine how boring it was to delete all documents titled as 'Untitled document'
among thousands of files under hundreds of folders. DriveWorks is a workspace 
addon to make file or folder management easy. The addon supports operations like 
deletion, renaming, etc on files/folders with filters like name matching, file 
type matching etc.  

# Design Choice
High level, the implementation of the addon used an async pattern, where user commits the job
spec into property storage first, then later, in the logic of triggers, the user job is executed 
and status is reported. Note that executing user jobs immediately after user committed
the job spec (inline execution) is only acceptable for small jobs and is going to be 
problematic (Google has limitation of each executation to be under 5 minutes) for jobs 
that lasts > 5 minutes or hours. 

To overcome the 5 minutes per executation limitation, the addon set up multiple triggers with 
5 minutes interval. At the end of each executation, we use property service to store the snapshot
of the executation. At the start of each executation, we use property service to resume the last 
executation state. Note that the trigger might not fire exactly at the configured time due to 
Google trigger limitations. However, setting up multiple triggers mitigate this problem to some degrees
and is acceptable in practice. 

# How To Use This Addon
## Step1. Clone the repo locally
`git clone https://github.com/ycui1984/DriveWorks.git`

## Step2. Create new projects for app scripts 
Open `https://script.google.com/home/start`
Click `New Project`

## Step3. Move the local files into the created project

## Step4. After you have the project
Click `Deploy` and `Test Deployments`

## Step5. Goto your google drive. The driveworks addon logo should show up at very right.

# ScreenShot
## Main Menu
![alt text](https://github.com/ycui1984/DriveWorks/blob/main/core/images/entry.png?raw=true)

## Delete File
![alt text](https://github.com/ycui1984/DriveWorks/blob/main/core/images/delete_file.png?raw=true)

## Delete Folder
![alt text](https://github.com/ycui1984/DriveWorks/blob/main/core/images/delete_folder.png?raw=true)

## Rename File
![alt text](https://github.com/ycui1984/DriveWorks/blob/main/core/images/rename_file.png?raw=true)

## Rename Folder
![alt text](https://github.com/ycui1984/DriveWorks/blob/main/core/images/rename_folder.png?raw=true)

# Demo Video
See https://www.youtube.com/watch?v=MWWHMgO7vP8 as an example on how to manage Google files/folders in batch.

# Donations
If this addon make your life easier, please consider to donate to encourage development.
Buymeacoffee https://www.buymeacoffee.com/smartgas
Paypal https://www.paypal.com/paypalme/SmartGAS888
