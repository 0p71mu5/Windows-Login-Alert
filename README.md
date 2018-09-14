# Windows-Login-Alert
Alerts user via text message on windows state change

It sends a text message alert to a pre-configured number whenever Windows System will be Logged In.
Uses way2sms account to send text message.

## Prerequisites:
- python 3 [Download](https://www.python.org/download/releases/3.0/)
- way2sms user account. [Register here](http://www.way2sms.com/user-registration)

## Installation:

After downloading python 3 run the below command in Windows-Login-Alert directory to install packages for Windows OS:
```
python -m pip install -r requirements.txt
```

## Config:
You need to mention Way2SMS registered mobile number and credentials along with alert receiving mobile number in __config.py__ file.
```
registered_mobile_windows_state_alert = ['9999999999']  # number registered on way2sms
receiver_number_windows_state_alert = ['9999999999']    # number where alert is to be sent
pass_windows_state_alert=['password']                   # credential for numbers registered in way2sms
```

Since, Way2SMS have limit of 25 text messages per day, you can create multiple accounts and mention them in __config.py__ file.
```
registered_mobile_windows_state_alert = ['9999999999','8888888888','7777777777']  # numbers registered on way2sms
receiver_number_windows_state_alert = ['9999999999']                              # number where alert is to be sent
pass_windows_state_alert=['password1','password2','password2']                    # credentials for numbers registered in way2sms
```

## Usage:
Copy the __windows_state_change_alert-login.bat__ file in 
C:\Users\\*username*\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup folder in Windows OS.

The script will call windows_state_change_alert.pyw file which will run in background.

To create alert for workstation lock and unlock state,
you need to create Tasks from __Task Scheduler__ and give condition as 'run a program' and manually mention scripts for each task.
- On workstation lock - windows_state_change_alert-locked.bat
- On workstation unlocked - windows_state_change_alert-unlocked.bat

A history of alerts will be maintained in __windows_state_info.xlsx__ file.
