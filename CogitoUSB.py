import subprocess
import keyboard
import time
import pandas as pd
import joblib
import sys
import re

#calculates words per minute typed
def wpm(chars, interval):
    return (chars/5)/interval

#calculates a list avg
def average(lst):
    len_check = len(lst)
    if len_check == 0:
        return 0
    else:
        return sum(lst)/len(lst)

#used to extract count of objects from powershell output
def run_int(cmd):
    p = subprocess.Popen(["powershell", "-Command", cmd], stdout=subprocess.PIPE)
    result = p.communicate()[0]
    return int(result)

#used to extract strings of object from a powershell output
def run_str(cmd):
    p = subprocess.Popen(["powershell", "-Command", cmd], stdout=subprocess.PIPE)
    result = p.communicate()[0]
    return str(result)

#Format devices in a list
def device_format(device_list):
    temp_list = re.findall("[VP]\w+", device_list)
    device_list = []
    dev_len = len(temp_list)
    device_list = [*set('\\'.join(x) for x in zip(temp_list[0::2], temp_list[1::2]))]
    return device_list

#detects USB Rubber Ducky via VendorID/ProductID and sets approriate flags
def ducky_detector1(usb_devices, hid_block, detected, blocked_devices, log_file):
    #check if blocked device is plugged in, if it is then detected is true
    for device in blocked_devices:
        if device in usb_devices:
            detected = True
    #if rubber ducky detected and device in block list, then HID is blocked
    if detected and not hid_block:
            print("----------------------------------------------------")
            print("USB Rubber Ducky Detected, HIDs disabled.")
            print(f"Blocked devices: {blocked_devices}")
            print("----------------------------------------------------\n")
            log_file.write("----------------------------------------------------\n")
            log_file.write("USB Rubber Ducky Detected, HIDs disabled.\n")
            log_file.write(f"Blocked devices: {blocked_devices}\n")
            log_file.write("----------------------------------------------------\n")
            hid_block = True
            return hid_block, detected, log_file
    #if no rubber ducky detected but HIDs still blocked, then unblock HIDs
    elif not detected and hid_block:
        print("----------------------------------------------------")
        print("USB Rubber Ducky has been removed, HIDs re-enabled.")
        print("----------------------------------------------------\n")
        log_file.write("----------------------------------------------------\n")
        log_file.write("USB Rubber Ducky has been removed, HIDs re-enabled.\n")
        log_file.write("----------------------------------------------------\n")
        hid_block = False
        keyboard.unhook_all()
        return hid_block, detected, log_file
    else:
        detected = False
        #check again if device in blocked list is detected
        for device in blocked_devices:
            if device in usb_devices:
                detected = True
        return hid_block, detected, log_file

def ducky_detector2(log_file):
    measuring = True
    time_tbl = []
    iki_tbl = []
    char_tbl = []
    word = ""

    event = keyboard.read_event() #wait for keyboard event
    t0 = time.time() #when event occurs, start timer

    #keep meaduring until space or enter keys are pressed, then end timer
    while measuring:
        event = keyboard.read_event()
        if 'down' in str(event):
            time_tbl.append(event.time)
            if event.name == 'space' or event.name == 'enter':
                char_tbl.append(' ')
                t1 = time.time()
                measuring = False
            else:
                char_tbl.append(event.name)

    wpm_speed = wpm(len(char_tbl), (t1-t0)/60) #num of characters divided by 5, over time interval in min
    for i in range(len(time_tbl)-1):
        iki_tbl.append((time_tbl[i+1] - time_tbl[i])*1000) #subtract elements from eachother and multiply by 1000 for miliseconds
    iki_speed = average(iki_tbl) #avg of above subtractions
    if iki_speed == 0:
        return ['human'], log_file

    data = pd.DataFrame([[wpm_speed, iki_speed]], columns=['avg_wpm', 'avg_iki']) #load data into data frame
    X = data
    model = joblib.load('cogitoUSB.joblib') #load cogitoUSB model
    prediction = model.predict(X) #used to predict if rubber ducky or human is typing based on above data and model
    print("----------------------------------------------------")
    print(f"Avg IKI: {iki_speed:.3f} ms")
    print(f"Avg WPM: {wpm_speed:.3f} wpm")
    print(f"Prediction: {prediction}")
    log_file.write("----------------------------------------------------\n")
    log_file.write(f"Avg IKI: {iki_speed:.3f} ms\n")
    log_file.write(f"Avg WPM: {wpm_speed:.3f} wpm\n")
    log_file.write(f"Prediction: {prediction}\n")
    log_file.write("----------------------------------------------------\n")
    return prediction, log_file

def hid_blocker(hid_block):
    if hid_block == True:
            for i in range(150):
                keyboard.block_key(i)

#used to quit the program, writes final blocked device list at the end of file before closing file
def quit_program(log_file, blocked_devices):
    log_file.write(f"\n{blocked_devices}\n")
    log_file.close()
    sys.exit(0)

def main():
    print("***************************************")
    print("** CogitoUSB - Rubber Ducky Detector **")
    print("***************************************\n")
    #used to determine of number of USB HIDs mounted to the system
    usb_enum_cnt = "(Get-PnpDevice -PresentOnly | Where-Object { $_.Class -match 'HIDClass' } |Where-Object {$_.InstanceId -match '^USB'}).Count"
    #used to determine of InstanceIds of USB HIDs mounted to the system
    usb_enum_devices = "(Get-PnpDevice -PresentOnly | Where-Object { $_.Class -match 'HIDClass' } |Where-Object {$_.InstanceId -match '^USB'}).InstanceID"
    #keeps track of initial number of USB HIDs mounted to the system at start of program
    usb_cnt_init = run_int(usb_enum_cnt)
    #Keeps track of initial HIDs by InstanceID
    usb_devices_init = run_str(usb_enum_devices)
    usb_devices_init = device_format(usb_devices_init)
    hid_block = False #default False to prevent needlessly blocking the keyboard
    detected = False #default False, as nothing has been detected yet
    ml_cadence_detector = []
    blocked_devices = []
    log_file = open("log.txt", "r+") #opens log_file to be appended
    lines = log_file.readlines() #reads log file by line into lines
    #searches from the end for most recently updated blocked devices list and assigns list to blocked_devices
    for line in reversed(lines):
        if line[0] == "[":
            temp_devices = line[1:-2]
            blocked_devices = temp_devices.replace("'", "").replace("\\\\", "\\").split(', ')
            break
    #if no blocked devices in log file, use default Rubber Ducky IDs as initial blocked devices list
    if blocked_devices == []:
        blocked_devices = ['VID_03EB\\PID_2401', 'VID_F000\\PID_FF02']
    print(f"Blocked list initialized: {blocked_devices}\n")
    allowed_devices = [] #keeps track of new HIDs that are not categorized as a Rubber Ducky
    print("Listening for Rubber Ducky...\n")

    while(True):
        if keyboard.is_pressed('ctrl+shift+a'):
            quit_program(log_file, blocked_devices)
        usb_cnt = run_int(usb_enum_cnt) #keeps track of quantity of USB HIDs plugged in
        usb_devices = run_str(usb_enum_devices) #keeps track of details of USB HIDs plugged in
        usb_devices = device_format(usb_devices) #format device list
        #if Rubber Ducky amongst HIDs, block input to HIDs keyboard otherwise unblock any hooks
        hid_block, detected, log_file = ducky_detector1(usb_devices, hid_block, detected, blocked_devices, log_file)
        #blocks all keys of keyboard if hid_block is True
        hid_blocker(hid_block)
        #if number of HIDs increases, triggers machine learning model cadence detector
        if (usb_cnt > usb_cnt_init) and not hid_block:
            ml_cadence_detector, log_file = ducky_detector2(log_file)
            #if rubber ducky detected and device not in blocked list, then device is added to blocked list and then HID is blocked
            if ml_cadence_detector == 'rubber_ducky' and not detected:
                for device in usb_devices:
                    if device not in usb_devices_init:
                        if device not in blocked_devices:
                            if device not in allowed_devices:
                                blocked_devices.append(device)
                print("----------------------------------------------------")
                print("USB Rubber Ducky Detected, HIDs disabled.")
                print(f"Blocked devices: {blocked_devices}")
                print("----------------------------------------------------\n")
                log_file.write("----------------------------------------------------\n")
                log_file.write(f"Device(s) added to block list.\n")
                log_file.write(f"Blocked devices: {blocked_devices}\n")
                log_file.write("----------------------------------------------------\n")
                hid_block = True
                detected = True
            #if rubber ducky detected and device in block list, then HID is blocked
            elif ml_cadence_detector == 'rubber_ducky' and detected:
                print("----------------------------------------------------")
                print("USB Rubber Ducky Detected, HIDs disabled.")
                print(f"Blocked devices: {blocked_devices}")
                print("----------------------------------------------------\n")
                log_file.write("----------------------------------------------------\n")
                log_file.write("USB Rubber Ducky Detected, HIDs disabled.\n")
                log_file.write(f"Blocked devices: {blocked_devices}\n")
                log_file.write("----------------------------------------------------\n")
                hid_block = True
            #if new device not rubber ducky and not in block list then add to approved list
            elif ml_cadence_detector  == 'human':
                print("----------------------------------------------------")
                print("USB Rubber Ducky not detected by cadence detector.")
                print("----------------------------------------------------\n")
                log_file.write("----------------------------------------------------\n")
                log_file.write("USB Rubber Ducky not detected by cadence detector.\n")
                log_file.write("----------------------------------------------------\n")
                hid_block = False
                for device in usb_devices:
                    if device not in usb_devices_init:
                        if device not in blocked_devices:
                            device = device.replace("\\r\\n'", "")
                            allowed_devices.append(device)
        #if Rubber Ducky amongst HIDs, block input to HIDs keyboard otherwise unblock any hooks
        hid_block, detected, log_file = ducky_detector1(usb_devices, hid_block, detected, blocked_devices, log_file)
        #blocks all keys of keyboard if hid_block is True
        hid_blocker(hid_block)

if __name__ == '__main__':
    main()
